from __future__ import annotations

import sys
import traceback
from pathlib import Path
from typing import Optional

from PySide6.QtCore import QThread, QTimer, QUrl, Signal
from PySide6.QtGui import QCloseEvent, QDesktopServices, QFont
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from collector import (
    ExportStatus,
    IChunQiuCollector,
    LeaderboardExportResult,
    MatchBindingInfo,
    OperationCancelled,
)


class BindWorker(QThread):
    log = Signal(str)
    success = Signal(object)
    error = Signal(str)
    cancelled = Signal()

    def __init__(self, url: str, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.url = url

    def run(self) -> None:
        collector = IChunQiuCollector(
            log_callback=lambda m: self.log.emit(m),
            cancel_check=self.isInterruptionRequested,
        )
        try:
            result = collector.bind_match(self.url)
            self.success.emit(result)
        except OperationCancelled:
            self.cancelled.emit()
        except Exception as exc:  # noqa: BLE001
            self.error.emit(f"{exc}\n\n{traceback.format_exc()}")
        finally:
            collector.close()


class ExportWorker(QThread):
    log = Signal(str)
    success = Signal(object)
    error = Signal(str)
    cancelled = Signal()

    def __init__(
        self,
        bind_info: MatchBindingInfo,
        board_keys: list[str],
        output_dir: str,
        page_size: int,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.bind_info = bind_info
        self.board_keys = board_keys
        self.output_dir = output_dir
        self.page_size = page_size

    def run(self) -> None:
        collector = IChunQiuCollector(
            log_callback=lambda m: self.log.emit(m),
            cancel_check=self.isInterruptionRequested,
        )
        try:
            results: list[LeaderboardExportResult] = []
            for board_key in self.board_keys:
                if self.isInterruptionRequested():
                    raise OperationCancelled("任务已取消")
                try:
                    result = collector.export_board_xlsx(
                        bind=self.bind_info,
                        output_root=self.output_dir,
                        board_key=board_key,
                        page_size=self.page_size,
                        debug_mode=False,
                    )
                except OperationCancelled:
                    raise
                except Exception as exc:  # noqa: BLE001
                    detail = f"{type(exc).__name__}: {exc}\n\n{traceback.format_exc()}"
                    self.log.emit(f"[Worker] {board_key} 未捕获异常: {detail}")
                    results.append(
                        LeaderboardExportResult(
                            board_key=board_key,
                            board_name=IChunQiuCollector.BOARD_CONFIGS.get(board_key, {}).get("name", board_key),
                            status=ExportStatus.FAILED,
                            excel_path="",
                            total_rows=0,
                            total_pages=0,
                            endpoint=IChunQiuCollector.BOARD_CONFIGS.get(board_key, {}).get("endpoint", ""),
                            error_message=detail,
                        )
                    )
                    continue
                results.append(result)
                self.log.emit(
                    f"[Worker] {board_key} -> {result.status} ({result.total_rows} 条)"
                )
            if self.isInterruptionRequested():
                raise OperationCancelled("任务已取消")
            self.success.emit(results)
        except OperationCancelled:
            self.cancelled.emit()
        finally:
            collector.close()


class MainWindow(QMainWindow):
    BOARD_BUTTONS = [
        ("solved", "导出解题总榜"),
        ("solved_dynamic", "导出解题动态"),
        ("integral", "导出积分总榜"),
        ("question", "导出题目榜单"),
        ("single", "导出单人总榜"),
        ("category", "导出题型榜单"),
        ("team_info", "导出团队信息"),
    ]

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("i春秋CTF信息采集器    微信公众号：星宇Sec")
        self.resize(1280, 820)
        self.setMinimumSize(1080, 700)

        self.current_bind: Optional[MatchBindingInfo] = None
        self.current_worker: Optional[QThread] = None
        self._close_pending = False
        self.last_export_dir: str = str((Path.cwd() / "outputs").resolve())
        self.export_btns: dict[str, QPushButton] = {}

        self._build_ui()
        self._apply_style()

    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("centralRoot")
        self.setCentralWidget(root)
        outer = QVBoxLayout(root)
        outer.setContentsMargins(14, 12, 14, 12)
        outer.setSpacing(10)

        header = QHBoxLayout()
        title = QLabel("一键采集i春秋CTF比赛观赛榜单信息")
        title.setFont(QFont("Microsoft YaHei UI", 18, QFont.Weight.Bold))
        self.status_label = QLabel("状态：就绪")
        self.status_label.setObjectName("statusLabel")
        self.cancel_btn = QPushButton("取消任务")
        self.cancel_btn.setObjectName("cancelButton")
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self.on_cancel_clicked)
        header.addWidget(title)
        header.addStretch(1)
        header.addWidget(self.cancel_btn)
        header.addWidget(self.status_label)
        outer.addLayout(header)

        config_group = self._build_config_group()
        outer.addWidget(config_group)

        info_group = self._build_info_group()
        outer.addWidget(info_group)

        log_group = QGroupBox("运行日志")
        log_layout = QVBoxLayout(log_group)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setPlaceholderText("这里显示绑定与导出的过程日志。")
        log_layout.addWidget(self.log_text)
        outer.addWidget(log_group, stretch=1)

    def _build_config_group(self) -> QGroupBox:
        group = QGroupBox("采集配置")
        layout = QGridLayout(group)
        layout.setHorizontalSpacing(10)
        layout.setVerticalSpacing(8)

        self.url_edit = QLineEdit()
        self.url_edit.setPlaceholderText("输入比赛 URL，例如 https://match.ichunqiu.com/2026hmg")
        self.bind_btn = QPushButton("绑定比赛信息")
        self.bind_btn.clicked.connect(self.on_bind_clicked)

        self.output_edit = QLineEdit(self.last_export_dir)
        self.choose_output_btn = QPushButton("选择目录")
        self.choose_output_btn.clicked.connect(self.on_choose_output)
        self.open_dir_btn = QPushButton("打开目录")
        self.open_dir_btn.clicked.connect(self.open_output_dir)

        self.page_size_spin = QSpinBox()
        self.page_size_spin.setRange(10, 200)
        self.page_size_spin.setValue(50)
        self.page_size_spin.setSingleStep(10)
        self.export_btns.clear()
        board_bar = QGridLayout()
        board_bar.setHorizontalSpacing(8)
        board_bar.setVerticalSpacing(6)

        self.collect_all_btn = QPushButton("一键采集全部榜单")
        self.collect_all_btn.setObjectName("collectAllButton")
        self.collect_all_btn.setEnabled(False)
        self.collect_all_btn.clicked.connect(self.on_collect_all_clicked)
        board_bar.addWidget(self.collect_all_btn, 0, 0, 1, 4)

        for idx, (board_key, text) in enumerate(self.BOARD_BUTTONS):
            btn = QPushButton(text)
            btn.setEnabled(False)
            btn.clicked.connect(lambda _=False, key=board_key: self.on_export_clicked(key))
            self.export_btns[board_key] = btn
            board_bar.addWidget(btn, idx // 4 + 1, idx % 4)

        layout.addWidget(QLabel("比赛 URL"), 0, 0)
        layout.addWidget(self.url_edit, 0, 1)
        layout.addWidget(self.bind_btn, 0, 2)

        layout.addWidget(QLabel("输出目录"), 1, 0)
        layout.addWidget(self.output_edit, 1, 1)
        dir_bar = QHBoxLayout()
        dir_bar.setSpacing(6)
        dir_bar.addWidget(self.choose_output_btn)
        dir_bar.addWidget(self.open_dir_btn)
        layout.addLayout(dir_bar, 1, 2)

        layout.addWidget(QLabel("每页条数"), 2, 0)
        layout.addWidget(self.page_size_spin, 2, 1)

        layout.addWidget(QLabel("导出榜单"), 3, 0)
        layout.addLayout(board_bar, 3, 1, 1, 2)

        layout.setColumnStretch(1, 1)
        return group

    def _readonly_line(self) -> QLineEdit:
        line = QLineEdit()
        line.setReadOnly(True)
        line.setObjectName("readonlyLine")
        line.setMinimumHeight(32)
        return line

    def _build_info_group(self) -> QGroupBox:
        group = QGroupBox("比赛信息")
        layout = QGridLayout(group)
        layout.setHorizontalSpacing(10)
        layout.setVerticalSpacing(8)

        self.title_line = self._readonly_line()
        self.event_line = self._readonly_line()
        self.k_line = self._readonly_line()
        self.time_line = self._readonly_line()
        self.situation_line = self._readonly_line()
        self.problems_line = self._readonly_line()

        self.open_situation_btn = QPushButton("打开链接")
        self.open_situation_btn.clicked.connect(
            lambda: self._open_external_url(self.situation_line.text().strip())
        )
        self.open_situation_btn.setEnabled(False)

        self.open_problems_btn = QPushButton("打开链接")
        self.open_problems_btn.clicked.connect(
            lambda: self._open_external_url(self.problems_line.text().strip())
        )
        self.open_problems_btn.setEnabled(False)

        layout.addWidget(QLabel("比赛标题"), 0, 0)
        layout.addWidget(self.title_line, 0, 1, 1, 4)

        layout.addWidget(QLabel("比赛标识"), 1, 0)
        layout.addWidget(self.event_line, 1, 1)
        layout.addWidget(QLabel("赛事访问令牌"), 1, 2)
        layout.addWidget(self.k_line, 1, 3, 1, 2)

        layout.addWidget(QLabel("比赛时间"), 2, 0)
        layout.addWidget(self.time_line, 2, 1, 1, 4)

        layout.addWidget(QLabel("观赛页"), 3, 0)
        layout.addWidget(self.situation_line, 3, 1, 1, 3)
        layout.addWidget(self.open_situation_btn, 3, 4)

        layout.addWidget(QLabel("排行榜页"), 4, 0)
        layout.addWidget(self.problems_line, 4, 1, 1, 3)
        layout.addWidget(self.open_problems_btn, 4, 4)

        layout.setColumnStretch(1, 2)
        layout.setColumnStretch(3, 4)
        return group

    def _apply_style(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow, QWidget#centralRoot {
                background: #f4f7fb;
            }
            QWidget {
                color: #162036;
                font-family: "Microsoft YaHei UI";
                font-size: 13px;
            }
            QLabel {
                background: transparent;
            }
            QGroupBox {
                border: 1px solid #d4deee;
                border-radius: 10px;
                margin-top: 10px;
                background: #ffffff;
                padding-top: 10px;
                font-weight: 600;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 2px 6px;
                background: transparent;
                color: #1e2a46;
            }
            QLineEdit, QTextEdit, QSpinBox {
                border: 1px solid #c8d5ea;
                border-radius: 8px;
                background: #fbfdff;
                padding: 6px 8px;
            }
            QLineEdit:focus, QTextEdit:focus, QSpinBox:focus {
                border: 1px solid #2f6ef5;
            }
            QLineEdit#readonlyLine {
                background: #f6f9ff;
                color: #27354f;
            }
            QPushButton {
                border: none;
                border-radius: 8px;
                padding: 8px 12px;
                background: #2f6ef5;
                color: white;
                font-weight: 600;
            }
            QPushButton:hover {
                background: #255ad1;
            }
            QPushButton:disabled {
                background: #9db5e6;
                color: #f4f7fd;
            }
            QPushButton#collectAllButton {
                background: #2f8f6b;
            }
            QPushButton#collectAllButton:hover {
                background: #287c5d;
            }
            QPushButton#collectAllButton:disabled {
                background: #a8cabc;
                color: #f2f7f4;
            }
            QPushButton#cancelButton {
                background: #c74646;
            }
            QPushButton#cancelButton:hover {
                background: #a93636;
            }
            QPushButton#cancelButton:disabled {
                background: #d8b7b7;
                color: #f7f2f2;
            }
            QLabel#statusLabel {
                background: transparent;
                color: #27437c;
                font-weight: 600;
            }
            """
        )

    def append_log(self, text: str) -> None:
        self.log_text.append(text)

    def set_busy(self, busy: bool, status: str) -> None:
        self.status_label.setText(f"状态：{status}")
        self.cancel_btn.setEnabled(busy)
        self.bind_btn.setEnabled(not busy)
        self.collect_all_btn.setEnabled((not busy) and self.current_bind is not None)
        for btn in self.export_btns.values():
            btn.setEnabled((not busy) and self.current_bind is not None)
        self.choose_output_btn.setEnabled(not busy)
        self.open_dir_btn.setEnabled(not busy)
        self.url_edit.setEnabled(not busy)
        self.output_edit.setEnabled(not busy)
        self.page_size_spin.setEnabled(not busy)
        self.open_situation_btn.setEnabled((not busy) and bool(self.situation_line.text().strip()))
        self.open_problems_btn.setEnabled((not busy) and bool(self.problems_line.text().strip()))

    def _open_external_url(self, url: str) -> None:
        if not url:
            QMessageBox.warning(self, "提示", "当前链接为空。")
            return
        ok = QDesktopServices.openUrl(QUrl(url))
        if not ok:
            QMessageBox.warning(self, "提示", f"无法打开链接：{url}")

    def on_choose_output(self) -> None:
        selected = QFileDialog.getExistingDirectory(
            self,
            "选择输出目录",
            self.output_edit.text().strip() or str(Path.cwd()),
        )
        if selected:
            self.output_edit.setText(selected)

    def on_bind_clicked(self) -> None:
        url = self.url_edit.text().strip()
        if not url:
            QMessageBox.warning(self, "提示", "请输入比赛地址。")
            return

        self.append_log("=" * 96)
        self.append_log(f"开始绑定比赛: {url}")
        self.set_busy(True, "正在绑定")

        worker = BindWorker(url, self)
        self.current_worker = worker
        worker.log.connect(self.append_log)
        worker.success.connect(self.on_bind_success)
        worker.error.connect(self.on_worker_error)
        worker.cancelled.connect(self.on_worker_cancelled)
        worker.finished.connect(self.on_worker_finished)
        worker.finished.connect(worker.deleteLater)
        worker.start()

    def on_bind_success(self, bind_obj: object) -> None:
        bind_info = bind_obj if isinstance(bind_obj, MatchBindingInfo) else None
        if bind_info is None:
            self.on_worker_error("绑定返回对象异常。")
            return

        self.current_bind = bind_info
        self.title_line.setText(bind_info.title)
        self.event_line.setText(bind_info.event_key)
        self.k_line.setText(bind_info.k)
        self.time_line.setText(f"{bind_info.start_time}  ->  {bind_info.end_time}")
        self.situation_line.setText(bind_info.situation_url)
        self.problems_line.setText(bind_info.problems_url)

        self.append_log("绑定成功。")
        self.set_busy(False, "绑定成功")

    def _start_export(self, board_keys: list[str], title_text: str) -> None:
        if self.current_bind is None:
            QMessageBox.warning(self, "提示", "请先绑定比赛信息。")
            return

        out_dir = self.output_edit.text().strip()
        if not out_dir:
            QMessageBox.warning(self, "提示", "请选择输出目录。")
            return
        Path(out_dir).mkdir(parents=True, exist_ok=True)
        self.last_export_dir = out_dir

        self.append_log("-" * 96)
        self.append_log(f"开始{title_text}，输出目录: {out_dir}")
        self.set_busy(True, "正在导出")

        worker = ExportWorker(
            bind_info=self.current_bind,
            board_keys=board_keys,
            output_dir=out_dir,
            page_size=int(self.page_size_spin.value()),
            parent=self,
        )
        self.current_worker = worker
        worker.log.connect(self.append_log)
        worker.success.connect(self.on_export_success)
        worker.error.connect(self.on_worker_error)
        worker.cancelled.connect(self.on_worker_cancelled)
        worker.finished.connect(self.on_worker_finished)
        worker.finished.connect(worker.deleteLater)
        worker.start()

    def on_export_clicked(self, board_key: str) -> None:
        board_name = IChunQiuCollector.BOARD_CONFIGS.get(board_key, {}).get("name", board_key)
        self._start_export([board_key], f"导出{board_name}")

    def on_collect_all_clicked(self) -> None:
        board_keys = [key for key, _ in self.BOARD_BUTTONS]
        self._start_export(board_keys, "一键采集全部榜单")

    def on_cancel_clicked(self) -> None:
        worker = self.current_worker
        if worker is None or not worker.isRunning():
            return
        self.append_log("正在取消任务，请等待当前请求结束...")
        self.status_label.setText("状态：正在取消")
        self.cancel_btn.setEnabled(False)
        worker.requestInterruption()

    def on_worker_cancelled(self) -> None:
        self.append_log("任务已取消，未完成的文件已清理。")
        if not self._close_pending:
            self.set_busy(False, "已取消")

    def on_worker_finished(self) -> None:
        worker = self.sender()
        if worker is self.current_worker:
            self.current_worker = None
        if self._close_pending:
            QTimer.singleShot(0, self.close)

    def on_export_success(self, result_obj: object) -> None:
        if not isinstance(result_obj, list) or not all(
            isinstance(item, LeaderboardExportResult) for item in result_obj
        ):
            self.on_worker_error("导出返回对象异常。")
            return
        results: list[LeaderboardExportResult] = result_obj

        # 分桶
        success_list = [r for r in results if r.status == ExportStatus.SUCCESS]
        partial_list = [r for r in results if r.status == ExportStatus.PARTIAL_SUCCESS]
        no_data_list = [r for r in results if r.status == ExportStatus.NO_DATA]
        failed_list = [r for r in results if r.status == ExportStatus.FAILED]

        for result in results:
            self.append_log(f"[{result.status}] {result.board_name} ({result.board_key})")
            self.append_log(f"  接口        : {result.endpoint}")
            self.append_log(f"  总条数      : {result.total_rows}, 总页数: {result.total_pages}")
            if result.excel_path:
                self.append_log(f"  Excel       : {result.excel_path}")
            if result.error_message:
                self.append_log(f"  说明/错误   : {result.error_message}")

        if self._close_pending:
            return
        self.set_busy(False, "导出完成")

        # 单个导出仍走原路径
        if len(results) == 1:
            r = results[0]
            if r.status == ExportStatus.SUCCESS:
                QMessageBox.information(self, "完成", f"{r.board_name}导出成功。\n\n{r.excel_path}")
            elif r.status == ExportStatus.PARTIAL_SUCCESS:
                box = QMessageBox(self)
                box.setIcon(QMessageBox.Warning)
                box.setWindowTitle("部分完成")
                box.setText(f"{r.board_name}已导出，但部分数据获取失败。")
                box.setInformativeText(r.excel_path)
                box.setDetailedText(r.error_message)
                box.setStandardButtons(QMessageBox.Ok)
                box.exec()
            elif r.status == ExportStatus.NO_DATA:
                QMessageBox.information(
                    self, "完成", f"{r.board_name}暂无数据，未生成 Excel 文件。"
                )
            else:
                self._show_error_box(f"{r.board_name}导出失败", r.error_message or "未知错误")
            return

        # 多个榜单：弹汇总
        self._show_summary(results, success_list, partial_list, no_data_list, failed_list)

    def on_worker_error(self, detail: str) -> None:
        self.append_log("任务失败：")
        self.append_log(detail)
        if self._close_pending:
            return
        self.set_busy(False, "失败")
        self._show_error_box("任务执行失败", detail)

    def _show_error_box(self, title: str, detail: str) -> None:
        """统一的错误弹窗：主文本简短，详细区放完整 traceback，避免被 QMessageBox 截断。"""
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Critical)
        box.setWindowTitle(title)
        box.setText(title)
        # 截断主文本以避免撑爆窗口
        short = detail.split("\n", 1)[0][:300]
        box.setInformativeText(short)
        box.setDetailedText(detail)
        box.setStandardButtons(QMessageBox.Ok)
        box.exec()

    def _show_summary(
        self,
        all_results: list[LeaderboardExportResult],
        success_list: list[LeaderboardExportResult],
        partial_list: list[LeaderboardExportResult],
        no_data_list: list[LeaderboardExportResult],
        failed_list: list[LeaderboardExportResult],
    ) -> None:
        """一键采集后的汇总弹窗。"""
        s_count = len(success_list)
        p_count = len(partial_list)
        n_count = len(no_data_list)
        f_count = len(failed_list)
        total = len(all_results)

        # 拼主文本
        lines = [f"共处理 {total} 个榜单："]
        lines.append(f"  ✅ 成功 {s_count} 个")
        lines.append(f"  ⚠️  部分成功 {p_count} 个")
        lines.append(f"  ⚠️  无数据 {n_count} 个")
        lines.append(f"  ❌ 失败 {f_count} 个")

        if success_list:
            lines.append("")
            lines.append("成功导出的文件：")
            for r in success_list:
                lines.append(f"  • {r.board_name} ({r.total_rows} 行)")
                if r.excel_path:
                    lines.append(f"      {r.excel_path}")

        if no_data_list:
            lines.append("")
            lines.append("无数据的榜单：")
            for r in no_data_list:
                lines.append(f"  • {r.board_name} — {r.error_message or '暂无数据'}")
                if r.excel_path:
                    lines.append(f"      提示文件: {r.excel_path}")

        if partial_list:
            lines.append("")
            lines.append("部分成功的文件：")
            for r in partial_list:
                lines.append(f"  • {r.board_name} ({r.total_rows} 行)")
                if r.excel_path:
                    lines.append(f"      {r.excel_path}")

        # 部分成功和失败列表拼到详细区
        detail_lines: list[str] = []
        if partial_list:
            detail_lines.append("部分成功明细：")
            for r in partial_list:
                detail_lines.append("")
                detail_lines.append(f"--- {r.board_name} ({r.board_key}) ---")
                detail_lines.append(r.error_message or "部分数据获取失败")
        if failed_list:
            if detail_lines:
                detail_lines.append("")
            detail_lines.append("失败明细：")
            for r in failed_list:
                detail_lines.append("")
                detail_lines.append(f"--- {r.board_name} ({r.board_key}) ---")
                detail_lines.append(f"endpoint: {r.endpoint}")
                detail_lines.append(r.error_message or "未知错误")

        box = QMessageBox(self)
        box.setIcon(
            QMessageBox.Information
            if f_count == 0 and p_count == 0
            else QMessageBox.Warning
        )
        box.setWindowTitle("一键采集完成")
        box.setText("\n".join(lines))
        if detail_lines:
            box.setDetailedText("\n".join(detail_lines))
        box.setStandardButtons(QMessageBox.Ok)
        box.exec()

    def open_output_dir(self) -> None:
        target = self.output_edit.text().strip() or self.last_export_dir
        path = Path(target)
        if not path.exists():
            QMessageBox.warning(self, "提示", f"目录不存在: {path}")
            return
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(path.resolve()))):
            QMessageBox.warning(self, "提示", f"无法打开目录: {path}")

    def closeEvent(self, event: QCloseEvent) -> None:  # noqa: N802
        worker = self.current_worker
        if worker is None or not worker.isRunning():
            event.accept()
            return

        if not self._close_pending:
            answer = QMessageBox.question(
                self,
                "任务正在运行",
                "当前任务尚未完成。是否取消任务并在清理完成后退出？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                event.ignore()
                return
            self._close_pending = True
            self.append_log("正在取消任务，清理完成后将自动退出...")
            self.status_label.setText("状态：正在取消并退出")
            self.cancel_btn.setEnabled(False)
            worker.requestInterruption()

        event.ignore()


def main() -> None:
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
