# -*- coding: utf-8 -*-
#
# Copyright © keithleygui Project Contributors
# Licensed under the terms of the MIT License
# (see LICENSE.txt for details)

# system imports
import os.path as osp
import configparser as cp
import time

# external imports
import pkg_resources as pkgr
import pyvisa
from PyQt5 import QtCore, QtWidgets, uic
from keithley2600 import Keithley2600, FETResultTable
from keithley2600.keithley_driver import KeithleyIOError
import numpy as np

# local imports
from keithleygui.pyqt_labutils import LedIndicator, SettingsWidget, ConnectionDialog
from keithleygui.pyqtplot_canvas import SweepDataPlot
from keithleygui.config.main import CONF

MAIN_UI_PATH = pkgr.resource_filename("keithleygui", "main.ui")


def _get_smus(keithley):
    """安全获取SMU列表，如果无法获取，则返回默认值"""
    try:
        # 检查设备是否已连接
        if hasattr(keithley, 'connected') and not keithley.connected:
            # 设备未连接，返回默认SMU列表
            smu_list = ["smu1", "smu2"]
        else:
            # 尝试从设备获取SMU列表
            smu_list = [attr_name for attr_name in dir(keithley) if attr_name.startswith("smu")]
    except Exception:
        # 出现任何错误，返回默认值
        smu_list = ["smu1", "smu2"]

    # we need at least two SMUs for GUI
    # -> fill with placeholders if required
    while len(smu_list) < 2:
        smu_list.append("--")

    return smu_list


class SMUSettingsWidget(SettingsWidget):

    SENSE_LOCAL = 0
    SENSE_REMOTE = 1

    def __init__(self, smu_name):
        super().__init__()

        self.smu_name = smu_name

        self.sense_type = self.addSelectionField(
            "Sense type:", ["local (2-wire)", "remote (4-wire)"]
        )
        self.limit_i = self.addDoubleField("Current limit:", 0.1, "A", limits=[0, 100])
        self.limit_v = self.addDoubleField("Voltage limit:", 200, "V", limits=[0, 200])
        self.high_c = self.addCheckBox("High capacitance mode", checked=False)

        self.load_defaults()

    def load_defaults(self):

        if self.smu_name != "--":

            try:
                sense_mode = CONF.get(self.smu_name, "sense")

                if sense_mode == "SENSE_LOCAL":
                    self.sense_type.setCurrentIndex(self.SENSE_LOCAL)
                elif sense_mode == "SENSE_REMOTE":
                    self.sense_type.setCurrentIndex(self.SENSE_REMOTE)

                self.limit_i.setValue(CONF.get(self.smu_name, "limiti"))
                self.limit_v.setValue(CONF.get(self.smu_name, "limitv"))
                self.high_c.setChecked(CONF.get(self.smu_name, "highc"))
            except cp.NoSectionError:
                pass

    def save_defaults(self):

        if self.smu_name != "--":

            if self.sense_type.currentIndex() == self.SENSE_LOCAL:
                CONF.set(self.smu_name, "sense", "SENSE_LOCAL")
            elif self.sense_type.currentIndex() == self.SENSE_REMOTE:
                CONF.set(self.smu_name, "sense", "SENSE_REMOTE")

            CONF.set(self.smu_name, "limiti", self.limit_i.value())
            CONF.set(self.smu_name, "limitv", self.limit_v.value())
            CONF.set(self.smu_name, "highc", self.high_c.isChecked())


# noinspection PyArgumentList
class SweepSettingsWidget(SettingsWidget):
    def __init__(self, keithley):
        super().__init__()

        self.keithley = keithley
        self.smu_list = _get_smus(self.keithley)

        self.t_int = self.addDoubleField("Integration time:", 0.1, "s", [0.000016, 5.0])
        self.t_settling = self.addDoubleField(
            "Settling time (auto = -1):", -1, "s", [-1, 100]
        )
        self.sweep_type = self.addSelectionField(
            "Sweep type:", ["Continuous", "Pulsed"]
        )
        self.smu_gate = self.addSelectionField("Gate SMU:", self.smu_list, 0)
        self.smu_drain = self.addSelectionField("Drain SMU:", self.smu_list, 1)

        self.load_defaults()

        self.smu_gate.currentIndexChanged.connect(self.on_smu_gate_changed)
        self.smu_drain.currentIndexChanged.connect(self.on_smu_drain_changed)

    def update_smu_list(self):
        try:
            self.smu_list = _get_smus(self.keithley)
    
            self.smu_gate.clear()
            self.smu_drain.clear()
    
            self.smu_gate.addItems(self.smu_list)
            self.smu_drain.addItems(self.smu_list)
    
            self.smu_gate.setCurrentIndex(0)
            self.smu_drain.setCurrentIndex(1)
        except Exception as e:
            print(f"更新SweepSettingsWidget的SMU列表时出错: {str(e)}")
            
    def load_defaults(self):

        self.t_int.setValue(CONF.get("Sweep", "tInt"))
        self.t_settling.setValue(CONF.get("Sweep", "delay"))
        self.sweep_type.setCurrentIndex(int(CONF.get("Sweep", "pulsed")))
        self.smu_gate.setCurrentText(CONF.get("Sweep", "gate"))
        self.smu_drain.setCurrentText(CONF.get("Sweep", "drain"))

    def save_defaults(self):
        CONF.set("Sweep", "tInt", self.t_int.value())
        CONF.set("Sweep", "delay", self.t_settling.value())
        CONF.set("Sweep", "gate", self.smu_gate.currentText())
        CONF.set("Sweep", "drain", self.smu_drain.currentText())

    @QtCore.pyqtSlot(int)
    def on_smu_gate_changed(self, int_smu):
        """Triggered when the user selects a different gate SMU. """

        if int_smu == 0 and len(self.smu_list) < 3:
            self.smu_drain.setCurrentIndex(1)
        elif int_smu == 1 and len(self.smu_list) < 3:
            self.smu_drain.setCurrentIndex(0)

    @QtCore.pyqtSlot(int)
    def on_smu_drain_changed(self, int_smu):
        """Triggered when the user selects a different drain SMU. """

        if int_smu == 0 and len(self.smu_list) < 3:
            self.smu_gate.setCurrentIndex(1)
        elif int_smu == 1 and len(self.smu_list) < 3:
            self.smu_gate.setCurrentIndex(0)


class TransferSweepSettingsWidget(SettingsWidget):
    def __init__(self):
        super().__init__()

        self.vg_start = self.addDoubleField("Vg start:", 0, "V")
        self.vg_stop = self.addDoubleField("Vg stop:", 0, "V")
        self.vg_step = self.addDoubleField("Vg step:", 0, "V")
        self.vd_list = self.addListField("Drain voltages:", [-5, -60])
        self.vd_list.setAcceptedStrings(["trailing"])

        self.load_defaults()

    def load_defaults(self):

        self.vg_start.setValue(CONF.get("Sweep", "VgStart"))
        self.vg_stop.setValue(CONF.get("Sweep", "VgStop"))
        self.vg_step.setValue(CONF.get("Sweep", "VgStep"))
        self.vd_list.setValue(CONF.get("Sweep", "VdList"))

    def save_defaults(self):
        CONF.set("Sweep", "VgStart", self.vg_start.value())
        CONF.set("Sweep", "VgStop", self.vg_stop.value())
        CONF.set("Sweep", "VgStep", self.vg_step.value())
        CONF.set("Sweep", "VdList", self.vd_list.value())


class OutputSweepSettingsWidget(SettingsWidget):
    def __init__(self):
        super().__init__()

        self.vd_start = self.addDoubleField("Vd start:", 0, "V")
        self.vd_stop = self.addDoubleField("Vd stop:", 0, "V")
        self.vd_step = self.addDoubleField("Vd step:", 0, "V")
        self.vg_list = self.addListField("Gate voltages:", [0, -20, -40, -60])

        self.load_defaults()

    def load_defaults(self):

        self.vd_start.setValue(CONF.get("Sweep", "VdStart"))
        self.vd_stop.setValue(CONF.get("Sweep", "VdStop"))
        self.vd_step.setValue(CONF.get("Sweep", "VdStep"))
        self.vg_list.setValue(CONF.get("Sweep", "VgList"))

    def save_defaults(self):
        CONF.set("Sweep", "VdStart", self.vd_start.value())
        CONF.set("Sweep", "VdStop", self.vd_stop.value())
        CONF.set("Sweep", "VdStep", self.vd_step.value())
        CONF.set("Sweep", "VgList", self.vg_list.value())


class IVSweepSettingsWidget(SettingsWidget):
    def __init__(self, keithley):
        super().__init__()
        self.keithley = keithley
        try:
            self.smu_list = _get_smus(self.keithley)
        except Exception:
            self.smu_list = ["smu1", "smu2"]

        self.v_start = self.addDoubleField("Vd start:", 0, "V")
        self.v_stop = self.addDoubleField("Vd stop:", 0, "V")
        self.v_step = self.addDoubleField("Vd step:", 0, "V")
        self.smu_sweep = self.addSelectionField("Sweep SMU:", self.smu_list, 0)

        self.load_defaults()

    def update_smu_list(self):
        try:
            self.smu_list = _get_smus(self.keithley)
            self.smu_sweep.clear()
            self.smu_sweep.addItems(self.smu_list)
            self.smu_sweep.setCurrentIndex(0)
        except Exception as e:
            print(f"更新IVSweepSettingsWidget的SMU列表时出错: {str(e)}")

    def load_defaults(self):

        self.v_start.setValue(CONF.get("Sweep", "VStart"))
        self.v_stop.setValue(CONF.get("Sweep", "VStop"))
        self.v_step.setValue(CONF.get("Sweep", "VStep"))
        self.smu_sweep.setCurrentText(CONF.get("Sweep", "smu_sweep"))

    def save_defaults(self):
        CONF.set("Sweep", "VStart", self.v_start.value())
        CONF.set("Sweep", "VStop", self.v_stop.value())
        CONF.set("Sweep", "VStep", self.v_step.value())
        CONF.set("Sweep", "smu_sweep", self.smu_sweep.currentText())


# noinspection PyArgumentList
class KeithleyGuiApp(QtWidgets.QMainWindow):
    """ Provides a GUI for transfer and output sweeps on the Keithley 2600."""

    QUIT_ON_CLOSE = True

    def __init__(self, keithley=None):
        super().__init__()
        # load user interface layout from .ui file
        uic.loadUi(MAIN_UI_PATH, self)

        try:
            if keithley:
                self.keithley = keithley
            else:
                address = CONF.get("Connection", "VISA_ADDRESS")
                lib = CONF.get("Connection", "VISA_LIBRARY")
                # 初始化Keithley对象但不立即尝试连接
                self.keithley = Keithley2600(address, lib)
                # 将connected属性设为False以避免尝试通信
                self.keithley.connected = False
            
            # 使用安全的方式获取SMU列表
            try:
                self.smu_list = _get_smus(self.keithley)
            except Exception:
                # 如果获取SMU列表失败，使用默认值
                self.smu_list = ["smu1", "smu2"]
                
        except Exception as e:
            # 出现任何错误时，使用默认值并显示错误消息
            self.keithley = Keithley2600("TCPIP0::localhost::0::SOCKET", "")
            self.keithley.connected = False
            self.smu_list = ["smu1", "smu2"]
            QtWidgets.QMessageBox.warning(
                self, 
                "初始化错误", 
                f"初始化Keithley设备时出错: {str(e)}\n\n程序将以离线模式运行。"
            )
            
        self.sweep_data = None

        # create sweep settings panes
        self.transfer_sweep_settings = TransferSweepSettingsWidget()
        self.output_sweep_settings = OutputSweepSettingsWidget()
        self.iv_sweep_settings = IVSweepSettingsWidget(self.keithley)
        self.general_sweep_settings = SweepSettingsWidget(self.keithley)

        self.tabWidgetSweeps.widget(0).layout().addWidget(self.transfer_sweep_settings)
        self.tabWidgetSweeps.widget(1).layout().addWidget(self.output_sweep_settings)
        self.tabWidgetSweeps.widget(2).layout().addWidget(self.iv_sweep_settings)
        self.groupBoxSweepSettings.layout().addWidget(self.general_sweep_settings)

        # create tabs for smu settings
        self.smu_tabs = []
        for smu_name in self.smu_list:
            tab = SMUSettingsWidget(smu_name)
            self.tabWidgetSettings.addTab(tab, smu_name)
            self.smu_tabs.append(tab)

        # create plot widget
        self.canvas = SweepDataPlot()
        self.gridLayout2.addWidget(self.canvas)

        # create LED indicator
        self.led = LedIndicator(self)
        self.statusBar.addPermanentWidget(self.led)
        self.led.setChecked(False)

        # create connection dialog
        self.connectionDialog = ConnectionDialog(self, self.keithley, CONF)

        # 添加实时测量标签页
        self.realtime_tab = QtWidgets.QWidget()
        self.tabWidgetSweeps.addTab(self.realtime_tab, "实时测量")
        realtime_layout = QtWidgets.QVBoxLayout(self.realtime_tab)
        
        # 添加实时测量控件
        realtime_form = QtWidgets.QFormLayout()
        self.realtime_smu = QtWidgets.QComboBox()
        self.realtime_smu.addItems(self.smu_list)
        self.realtime_interval = QtWidgets.QDoubleSpinBox()
        self.realtime_interval.setRange(0.1, 10.0)
        self.realtime_interval.setValue(0.5)
        self.realtime_interval.setSuffix(" s")
        
        # 电压控制部分
        voltage_control_layout = QtWidgets.QHBoxLayout()
        self.realtime_voltage = QtWidgets.QDoubleSpinBox()
        self.realtime_voltage.setRange(-200, 200)
        self.realtime_voltage.setValue(0)
        self.realtime_voltage.setSuffix(" V")
        self.realtime_voltage.setSingleStep(0.1)  # 步进值为0.1V
        
        # 添加快速电压增减按钮
        self.voltage_up_button = QtWidgets.QPushButton("+1V")
        self.voltage_down_button = QtWidgets.QPushButton("-1V")
        self.voltage_zero_button = QtWidgets.QPushButton("0V")
        
        voltage_buttons_layout = QtWidgets.QHBoxLayout()
        voltage_buttons_layout.addWidget(self.voltage_down_button)
        voltage_buttons_layout.addWidget(self.voltage_zero_button)
        voltage_buttons_layout.addWidget(self.voltage_up_button)
        
        voltage_control_layout.addWidget(self.realtime_voltage)
        
        realtime_form.addRow("测量SMU:", self.realtime_smu)
        realtime_form.addRow("更新间隔:", self.realtime_interval)
        realtime_form.addRow("应用电压:", voltage_control_layout)
        realtime_form.addRow("快速调节:", voltage_buttons_layout)

        # 添加实时测量值显示
        self.realtime_values = QtWidgets.QGroupBox("实时测量值")
        realtime_values_layout = QtWidgets.QFormLayout(self.realtime_values)
        
        self.realtime_time_label = QtWidgets.QLabel("0.00 s")
        self.realtime_voltage_label = QtWidgets.QLabel("0.00 V")
        self.realtime_current_label = QtWidgets.QLabel("0.00 A")
        
        realtime_values_layout.addRow("时间:", self.realtime_time_label)
        realtime_values_layout.addRow("电压:", self.realtime_voltage_label)
        realtime_values_layout.addRow("电流:", self.realtime_current_label)
        
        # 添加到主布局
        realtime_layout.addLayout(realtime_form)
        realtime_layout.addWidget(self.realtime_values)

        # 添加实时测量按钮
        realtime_buttons = QtWidgets.QHBoxLayout()
        self.pushButtonStartRealtime = QtWidgets.QPushButton("开始实时测量")
        self.pushButtonStopRealtime = QtWidgets.QPushButton("停止实时测量")
        self.pushButtonStopRealtime.setEnabled(False)
        
        realtime_buttons.addWidget(self.pushButtonStartRealtime)
        realtime_buttons.addWidget(self.pushButtonStopRealtime)
        realtime_layout.addLayout(realtime_buttons)
        
        # 实时测量线程
        self.realtime_thread = None

        # restore last position and size
        self.restore_geometry()

        # update GUI status and connect callbacks
        self.actionSaveSweepData.setEnabled(False)
        self.connect_ui_callbacks()
        self.on_load_default()
        self.update_gui_connection()

        # connection update timer: check periodically if keithley is connected
        self.connection_status_update = QtCore.QTimer()
        self.connection_status_update.timeout.connect(self.update_gui_connection)
        self.connection_status_update.start(10000)  # 10 sec

    def update_smu_list(self):
        """Update all smu lists in the interface."""
        try:
            # 安全获取SMU列表
            self.smu_list = _get_smus(self.keithley)
    
            # update smu lists in widgets
            try:
                if hasattr(self, 'transfer_sweep_settings'):
                    self.transfer_sweep_settings.update_smu_list()
                if hasattr(self, 'output_sweep_settings'):
                    self.output_sweep_settings.update_smu_list()
                if hasattr(self, 'iv_sweep_settings'):
                    self.iv_sweep_settings.update_smu_list()
                if hasattr(self, 'general_sweep_settings'):
                    self.general_sweep_settings.update_smu_list()
            except Exception as e:
                print(f"更新SMU下拉列表出错: {str(e)}")
            
            # 更新实时测量SMU列表
            if hasattr(self, 'realtime_smu'):
                try:
                    current_text = self.realtime_smu.currentText()
                    self.realtime_smu.clear()
                    self.realtime_smu.addItems(self.smu_list)
                    # 尝试恢复之前的选择
                    index = self.realtime_smu.findText(current_text)
                    if index >= 0:
                        self.realtime_smu.setCurrentIndex(index)
                except Exception as e:
                    print(f"更新实时测量SMU列表出错: {str(e)}")
    
            # update smu settings tabs
            try:
                self.smu_tabs = []
                self.tabWidgetSettings.clear()
                for smu_name in self.smu_list:
                    tab = SMUSettingsWidget(smu_name)
                    self.tabWidgetSettings.addTab(tab, smu_name)
                    self.smu_tabs.append(tab)
                    
                if hasattr(self, 'iv_sweep_settings'):
                    self.iv_sweep_settings.update_smu_list()
                if hasattr(self, 'general_sweep_settings'):
                    self.general_sweep_settings.update_smu_list()
            except Exception as e:
                print(f"更新SMU设置标签页出错: {str(e)}")
                
        except Exception as e:
            print(f"更新SMU列表时出错: {str(e)}")

    @staticmethod
    def _string_to_vd(string):
        try:
            return float(string)
        except ValueError:
            if "trailing" in string:
                return "trailing"
            else:
                raise ValueError("Invalid drain voltage.")

    def closeEvent(self, event):
        if self.QUIT_ON_CLOSE:
            self.exit_()
        else:
            self.hide()

    # =============================================================================
    # GUI setup
    # =============================================================================

    def restore_geometry(self):
        x = CONF.get("Window", "x")
        y = CONF.get("Window", "y")
        w = CONF.get("Window", "width")
        h = CONF.get("Window", "height")

        self.setGeometry(x, y, w, h)

    def save_geometry(self):
        geo = self.geometry()
        CONF.set("Window", "height", geo.height())
        CONF.set("Window", "width", geo.width())
        CONF.set("Window", "x", geo.x())
        CONF.set("Window", "y", geo.y())

    def connect_ui_callbacks(self):
        """Connect buttons and menus to callbacks."""
        self.pushButtonRun.clicked.connect(self.on_sweep_clicked)
        self.pushButtonAbort.clicked.connect(self.on_abort_clicked)

        self.actionSettings.triggered.connect(self.connectionDialog.open)
        self.actionConnect.triggered.connect(self.on_connect_clicked)
        self.actionDisconnect.triggered.connect(self.on_disconnect_clicked)
        self.actionExit.triggered.connect(self.exit_)
        self.actionSaveSweepData.triggered.connect(self.on_save_clicked)
        self.actionLoad_data_from_file.triggered.connect(self.on_load_clicked)
        self.actionSaveDefaults.triggered.connect(self.on_save_default)
        self.actionLoadDefaults.triggered.connect(self.on_load_default)
        
        # 实时测量控件绑定
        self.pushButtonStartRealtime.clicked.connect(self.on_start_realtime_clicked)
        self.pushButtonStopRealtime.clicked.connect(self.on_stop_realtime_clicked)
        
        # 电压调节按钮绑定
        self.voltage_up_button.clicked.connect(self.on_voltage_up_clicked)
        self.voltage_down_button.clicked.connect(self.on_voltage_down_clicked)
        self.voltage_zero_button.clicked.connect(self.on_voltage_zero_clicked)

    # =============================================================================
    # Measurement callbacks
    # =============================================================================

    def apply_smu_settings(self):
        """
        Applies SMU settings to Keithley before a measurement.
        Warning: self.keithley.reset() will reset those settings.
        """
        for tab in self.smu_tabs:

            smu = getattr(self.keithley, tab.smu_name)

            if tab.sense_type.currentIndex() == tab.SENSE_LOCAL:
                smu.sense = smu.SENSE_LOCAL
            elif tab.sense_type.currentIndex() == tab.SENSE_REMOTE:
                smu.sense = smu.SENSE_REMOTE

            lim_i = tab.limit_i.value()
            smu.source.limiti = lim_i
            smu.trigger.source.limiti = lim_i

            lim_v = tab.limit_v.value()
            smu.source.limitv = lim_v
            smu.trigger.source.limitv = lim_v

            smu.source.highc = int(tab.high_c.isChecked())

    @QtCore.pyqtSlot()
    def on_sweep_clicked(self):
        """ Start a transfer measurement with current settings."""

        if self.keithley.busy:
            msg = "Keithley is currently busy. Please try again later."
            QtWidgets.QMessageBox.information(self, "Keithley Busy", msg)

            return

        self.apply_smu_settings()

        params = dict()

        if self.tabWidgetSweeps.currentIndex() == 0:
            self.statusBar.showMessage("    Recording transfer curve.")
            # get sweep settings
            params["sweep_type"] = "transfer"
            params["VgStart"] = self.transfer_sweep_settings.vg_start.value()
            params["VgStop"] = self.transfer_sweep_settings.vg_stop.value()
            params["VgStep"] = self.transfer_sweep_settings.vg_step.value()
            params["VdList"] = self.transfer_sweep_settings.vd_list.value()

        elif self.tabWidgetSweeps.currentIndex() == 1:
            self.statusBar.showMessage("    Recording output curve.")
            # get sweep settings
            params["sweep_type"] = "output"
            params["VdStart"] = self.output_sweep_settings.vd_start.value()
            params["VdStop"] = self.output_sweep_settings.vd_stop.value()
            params["VdStep"] = self.output_sweep_settings.vd_step.value()
            params["VgList"] = self.output_sweep_settings.vg_list.value()

        elif self.tabWidgetSweeps.currentIndex() == 2:
            self.statusBar.showMessage("    Recording IV curve.")
            # get sweep settings
            params["sweep_type"] = "iv"
            params["VStart"] = self.iv_sweep_settings.v_start.value()
            params["VStop"] = self.iv_sweep_settings.v_stop.value()
            params["VStep"] = self.iv_sweep_settings.v_step.value()
            smusweep = self.iv_sweep_settings.smu_sweep.currentText()
            params["smu_sweep"] = getattr(self.keithley, smusweep)

        else:
            return

        # get general sweep settings
        smu_gate = self.general_sweep_settings.smu_gate.currentText()
        smu_drain = self.general_sweep_settings.smu_drain.currentText()
        params["tInt"] = self.general_sweep_settings.t_int.value()
        params["delay"] = self.general_sweep_settings.t_settling.value()
        params["smu_gate"] = getattr(self.keithley, smu_gate)
        params["smu_drain"] = getattr(self.keithley, smu_drain)
        params["pulsed"] = bool(self.general_sweep_settings.sweep_type.currentIndex())

        # check if integration time is valid, return otherwise
        freq = self.keithley.localnode.linefreq

        if not 0.001 / freq < params["tInt"] < 25.0 / freq:
            msg = (
                "Integration time must be between 0.001 and 25 "
                + "power line cycles of 1/(%s Hz)." % freq
            )
            QtWidgets.QMessageBox.information(self, "Parameter Error", msg)

            return

        # create measurement thread with params dictionary
        self.measureThread = MeasureThread(self.keithley, params)
        self.measureThread.finished_sig.connect(self.on_measure_done)
        self.measureThread.error_sig.connect(self.on_measure_error)

        # run measurement
        self._gui_state_busy()
        self.measureThread.start()

    def on_measure_done(self, sd):
        self.statusBar.showMessage("    Ready.")
        self._gui_state_idle()
        self.actionSaveSweepData.setEnabled(True)

        self.sweep_data = sd
        self.canvas.plot(self.sweep_data)
        if not self.keithley.abort_event.is_set():
            self.on_save_clicked()

    def on_measure_error(self, exc):
        self.statusBar.showMessage("    Ready.")
        self._gui_state_idle()
        QtWidgets.QMessageBox.information(
            self, "Sweep Error", f"{exc.__class__.__name__}: {exc.args[0]}"
        )

    @QtCore.pyqtSlot()
    def on_abort_clicked(self):
        """
        Aborts current measurement.
        """
        self.keithley.abort_event.set()
        for smu in self.smu_list:
            getattr(self.keithley, smu).abort()
        self.keithley.reset()

    # =============================================================================
    # Interface callbacks
    # =============================================================================

    @QtCore.pyqtSlot()
    def on_connect_clicked(self):
        try:
            self.keithley.connect()
            self.update_smu_list()
            self.update_gui_connection()
            if not self.keithley.connected:
                msg = (
                    f"Keithley无法在{self.keithley.visa_address}地址连接。 "
                    f"请检查地址是否正确，Keithley设备是否已打开。"
                )
                QtWidgets.QMessageBox.information(self, "连接错误", msg)
        except Exception as e:
            self.keithley.connected = False
            msg = f"连接Keithley设备时出错:\n{str(e)}"
            QtWidgets.QMessageBox.information(self, "连接错误", msg)
            self.update_gui_connection()

    @QtCore.pyqtSlot()
    def on_disconnect_clicked(self):
        self.keithley.disconnect()
        self.update_gui_connection()
        self.statusBar.showMessage("    No Keithley connected.")

    @QtCore.pyqtSlot()
    def on_save_clicked(self):
        """Show GUI to save current sweep data as text file."""
        prompt = "Save as .txt file."
        filename = "untitled.txt"
        formats = "Text file (*.txt)"
        filepath, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, prompt, filename, formats
        )
        if len(filepath) < 4:
            return
        self.sweep_data.save(filepath)

    @QtCore.pyqtSlot()
    def on_load_clicked(self):
        """Show GUI to load sweep data from file."""
        prompt = "Please select a data file."
        filepath, _ = QtWidgets.QFileDialog.getOpenFileName(self, prompt)
        if not osp.isfile(filepath):
            return

        self.sweep_data = FETResultTable()
        self.sweep_data.load(filepath)

        self.canvas.plot(self.sweep_data)
        self.actionSaveSweepData.setEnabled(True)

    @QtCore.pyqtSlot()
    def on_save_default(self):
        """Saves current settings from GUI as defaults."""

        # save sweep settings
        self.transfer_sweep_settings.save_defaults()
        self.output_sweep_settings.save_defaults()
        self.iv_sweep_settings.save_defaults()
        self.general_sweep_settings.save_defaults()

        # save smu specific settings
        for tab in self.smu_tabs:
            tab.save_defaults()

    @QtCore.pyqtSlot()
    def on_load_default(self):
        """Load default settings to interface."""

        # load sweep settings
        self.transfer_sweep_settings.load_defaults()
        self.output_sweep_settings.load_defaults()
        self.iv_sweep_settings.load_defaults()
        self.general_sweep_settings.load_defaults()

        # smu settings
        for tab in self.smu_tabs:
            tab.load_defaults()

    @QtCore.pyqtSlot()
    def exit_(self):
        self.keithley.disconnect()
        self.connection_status_update.stop()
        self.save_geometry()
        self.deleteLater()

    # =============================================================================
    # Interface states
    # =============================================================================

    def update_gui_connection(self):
        """Check if Keithley is connected and update GUI."""
        try:
            if not hasattr(self.keithley, 'connected'):
                self.keithley.connected = False
                self._gui_state_disconnected()
                return
                
            if self.keithley.connected:
                try:
                    test = self.keithley.localnode.model
                except (
                    pyvisa.VisaIOError,
                    pyvisa.InvalidSession,
                    OSError,
                    KeithleyIOError,
                    AttributeError,
                    TypeError,
                    Exception,
                ):
                    self.keithley.connected = False
                    self._gui_state_disconnected()
                else:
                    if self.keithley.busy:
                        self._gui_state_busy()
                    else:
                        self._gui_state_idle()
            else:
                self._gui_state_disconnected()
        except Exception:
            # 发生任何错误，设置为断开连接状态
            if hasattr(self.keithley, 'connected'):
                self.keithley.connected = False
            self._gui_state_disconnected()

    def _gui_state_busy(self):
        """Set GUI to state for running measurement."""

        self.pushButtonRun.setEnabled(False)
        self.pushButtonAbort.setEnabled(True)

        self.actionConnect.setEnabled(False)
        self.actionDisconnect.setEnabled(False)

        self.statusBar.showMessage("    Measuring.")
        self.led.setChecked(True)

    def _gui_state_idle(self):
        """Set GUI to state for IDLE Keithley."""

        self.pushButtonRun.setEnabled(True)
        self.pushButtonAbort.setEnabled(False)

        self.actionConnect.setEnabled(False)
        self.actionDisconnect.setEnabled(True)
        self.statusBar.showMessage("    Ready.")
        self.led.setChecked(True)

    def _gui_state_disconnected(self):
        """ UI changes when keithley is disconnected."""
        self.actionConnect.setEnabled(True)
        self.actionDisconnect.setEnabled(False)
        
        self.pushButtonRun.setEnabled(False)
        self.pushButtonAbort.setEnabled(False)
        
        # 允许在模拟模式下使用实时测量
        self.pushButtonStartRealtime.setEnabled(True)
        self.pushButtonStopRealtime.setEnabled(False)
        
        if hasattr(self, 'apply_keithley_settings'):
            self.apply_keithley_settings.setEnabled(False)
            
        self.statusBar.showMessage("    No Keithley connected.")
        self.led.setChecked(False)

    @QtCore.pyqtSlot()
    def on_start_realtime_clicked(self):
        """开始实时测量"""
        # 获取选择的SMU和更新间隔
        smu_name = self.realtime_smu.currentText()
        interval = self.realtime_interval.value()
        voltage = self.realtime_voltage.value()
        
        if smu_name == "--":
            QtWidgets.QMessageBox.information(
                self, "参数错误", "请选择有效的SMU"
            )
            return
            
        # 检查是否是模拟模式
        simulation_mode = not self.keithley.connected
        
        if not simulation_mode:
            # 真实设备模式
            try:
                # 应用电压
                smu = getattr(self.keithley, smu_name)
                self.keithley.apply_voltage(smu, voltage)
            except Exception as e:
                QtWidgets.QMessageBox.information(
                    self, "设置错误", f"设置电压时出错: {str(e)}"
                )
                return
        else:
            # 模拟模式，显示提示信息
            msg_box = QtWidgets.QMessageBox()
            msg_box.setIcon(QtWidgets.QMessageBox.Information)
            msg_box.setWindowTitle("模拟模式")
            msg_box.setText("设备未连接，将使用模拟数据")
            msg_box.setStandardButtons(QtWidgets.QMessageBox.Ok | QtWidgets.QMessageBox.Cancel)
            
            if msg_box.exec_() == QtWidgets.QMessageBox.Cancel:
                return
        
        # 创建并启动实时测量线程
        self.realtime_thread = RealtimeMeasureThread(self.keithley, smu_name, interval, simulation_mode)
        self.realtime_thread.data_sig.connect(self.on_realtime_data)
        self.realtime_thread.error_sig.connect(self.on_realtime_error)
        self.realtime_thread.finished_sig.connect(self.on_realtime_finished)
        
        # 更新UI状态
        self.pushButtonStartRealtime.setEnabled(False)
        self.pushButtonStopRealtime.setEnabled(True)
        self.realtime_smu.setEnabled(False)
        self.realtime_interval.setEnabled(False)
        
        # 在测量过程中保持电压控制可用，并添加电压改变时的回调函数
        self.realtime_voltage.setEnabled(True)
        if not hasattr(self, 'voltage_value_changed_connected'):
            self.realtime_voltage.valueChanged.connect(self.on_voltage_changed)
            self.voltage_value_changed_connected = True
        
        # 清除之前的图表数据
        self.canvas.clear()
        
        # 启动线程
        self.realtime_thread.start()
        
        if simulation_mode:
            self.statusBar.showMessage("    正在进行模拟测量...")
        else:
            self.statusBar.showMessage("    正在进行实时测量...")
            
    @QtCore.pyqtSlot(float)
    def on_voltage_changed(self, new_voltage):
        """当电压控制值改变时调用"""
        if (hasattr(self, 'realtime_thread') and 
            self.realtime_thread and 
            self.realtime_thread.isRunning()):
            
            simulation_mode = hasattr(self.realtime_thread, 'simulation_mode') and self.realtime_thread.simulation_mode
            
            if not simulation_mode and self.keithley.connected:
                try:
                    smu_name = self.realtime_smu.currentText()
                    if smu_name != "--":
                        smu = getattr(self.keithley, smu_name)
                        # 应用新电压
                        self.keithley.apply_voltage(smu, new_voltage)
                        # 更新状态栏
                        self.statusBar.showMessage(f"    电压已更新: {new_voltage} V")
                except Exception as e:
                    print(f"更新电压时出错: {str(e)}")
                    self.statusBar.showMessage(f"    电压更新失败: {str(e)}")
            elif simulation_mode:
                # 在模拟模式下，更新模拟电压值
                if hasattr(self.realtime_thread, 'update_simulation_voltage'):
                    self.realtime_thread.update_simulation_voltage(new_voltage)
                    self.statusBar.showMessage(f"    模拟电压已更新: {new_voltage} V")

    def _reset_realtime_ui(self):
        """重置实时测量UI状态"""
        self.pushButtonStartRealtime.setEnabled(True)
        self.pushButtonStopRealtime.setEnabled(False)
        self.realtime_smu.setEnabled(True)
        self.realtime_interval.setEnabled(True)
        self.realtime_voltage.setEnabled(True)

    @QtCore.pyqtSlot()
    def on_stop_realtime_clicked(self):
        """停止实时测量"""
        if self.realtime_thread and self.realtime_thread.isRunning():
            self.realtime_thread.stop()
            self.statusBar.showMessage("    正在停止实时测量...")
    
    @QtCore.pyqtSlot(object)
    def on_realtime_data(self, data):
        """处理实时测量数据"""
        try:
            # 更新界面显示的实时值
            t, v, i = data["last_value"]
            self.realtime_time_label.setText(f"{t:.2f} s")
            self.realtime_voltage_label.setText(f"{v:.6f} V")
            self.realtime_current_label.setText(f"{i:.6e} A")
            
            # 更新图表
            if len(data["time"]) > 1:  # 至少有两个点才能画图
                # 绘制时间-电流曲线
                self.canvas.clear()
                self.canvas.plot_xy(data["time"], data["current"], "时间 (s)", "电流 (A)", "实时电流")
        except Exception as e:
            print(f"处理实时数据时出错: {str(e)}")
    
    @QtCore.pyqtSlot(object)
    def on_realtime_error(self, exc):
        """处理实时测量错误"""
        try:
            self.statusBar.showMessage("    测量出错.")
            QtWidgets.QMessageBox.information(
                self, "实时测量错误", f"{exc.__class__.__name__}: {exc.args[0]}"
            )
            self._reset_realtime_ui()
        except Exception as e:
            print(f"处理实时测量错误时出错: {str(e)}")
            self.statusBar.showMessage("    测量出错.")
            self._reset_realtime_ui()
    
    @QtCore.pyqtSlot()
    def on_realtime_finished(self):
        """实时测量结束后的处理"""
        try:
            self.statusBar.showMessage("    Ready.")
            self._reset_realtime_ui()
            
            # 检查是否需要关闭设备输出
            if (hasattr(self, 'realtime_thread') and 
                hasattr(self.realtime_thread, 'simulation_mode') and 
                not self.realtime_thread.simulation_mode):
                
                # 只在实际设备模式下关闭输出
                if hasattr(self, 'keithley') and hasattr(self.keithley, 'connected') and self.keithley.connected:
                    try:
                        smu_name = self.realtime_smu.currentText()
                        if smu_name != "--":
                            smu = getattr(self.keithley, smu_name)
                            smu.source.output = smu.OUTPUT_OFF
                    except Exception as e:
                        print(f"关闭输出时出错: {str(e)}")
        except Exception as e:
            print(f"实时测量结束处理时出错: {str(e)}")
            try:
                self._reset_realtime_ui()
            except:
                pass

    @QtCore.pyqtSlot()
    def on_voltage_up_clicked(self):
        """增加电压1V"""
        current_value = self.realtime_voltage.value()
        self.realtime_voltage.setValue(current_value + 1.0)
        
    @QtCore.pyqtSlot()
    def on_voltage_down_clicked(self):
        """减少电压1V"""
        current_value = self.realtime_voltage.value()
        self.realtime_voltage.setValue(current_value - 1.0)
        
    @QtCore.pyqtSlot()
    def on_voltage_zero_clicked(self):
        """将电压设为0V"""
        self.realtime_voltage.setValue(0.0)


# noinspection PyUnresolvedReferences
class MeasureThread(QtCore.QThread):

    started_sig = QtCore.pyqtSignal()
    finished_sig = QtCore.pyqtSignal(object)
    error_sig = QtCore.pyqtSignal(object)

    def __init__(self, keithley, params):
        QtCore.QThread.__init__(self)
        self.keithley = keithley
        self.params = params

    def __del__(self):
        self.wait()

    def run(self):

        self.started_sig.emit()

        try:
            sweep_data = None

            if self.params["sweep_type"] == "transfer":
                sweep_data = self.keithley.transfer_measurement(
                    self.params["smu_gate"],
                    self.params["smu_drain"],
                    self.params["VgStart"],
                    self.params["VgStop"],
                    self.params["VgStep"],
                    self.params["VdList"],
                    self.params["tInt"],
                    self.params["delay"],
                    self.params["pulsed"],
                )
            elif self.params["sweep_type"] == "output":
                sweep_data = self.keithley.output_measurement(
                    self.params["smu_gate"],
                    self.params["smu_drain"],
                    self.params["VdStart"],
                    self.params["VdStop"],
                    self.params["VdStep"],
                    self.params["VgList"],
                    self.params["tInt"],
                    self.params["delay"],
                    self.params["pulsed"],
                )

            elif self.params["sweep_type"] == "iv":
                direction = np.sign(self.params["VStop"] - self.params["VStart"])
                stp = direction * abs(self.params["VStep"])

                # forward and reverse sweeps
                sweeplist = np.arange(
                    self.params["VStart"], self.params["VStop"] + stp, stp
                )
                sweeplist = np.append(sweeplist, np.flip(sweeplist))

                v, i = self.keithley.voltage_sweep_single_smu(
                    self.params["smu_sweep"],
                    sweeplist,
                    self.params["tInt"],
                    self.params["delay"],
                    self.params["pulsed"],
                )

                params = {
                    "sweep_type": "iv",
                    "t_int": self.params["tInt"],
                    "delay": self.params["delay"],
                    "pulsed": self.params["pulsed"],
                }

                sweep_data = FETResultTable(
                    column_titles=["Voltage", "Current"],
                    units=["V", "A"],
                    data=np.array([v, i]).transpose(),
                    params=params,
                )

            self.keithley.beeper.beep(0.3, 2400)
            self.keithley.reset()

            self.finished_sig.emit(sweep_data)
        except Exception as exc:
            self.error_sig.emit(exc)


# 实时测量线程
class RealtimeMeasureThread(QtCore.QThread):
    data_sig = QtCore.pyqtSignal(object)
    error_sig = QtCore.pyqtSignal(object)
    finished_sig = QtCore.pyqtSignal()

    def __init__(self, keithley, smu_name, interval=0.5, simulation_mode=False):
        QtCore.QThread.__init__(self)
        self.keithley = keithley
        self.smu_name = smu_name
        try:
            self.smu = getattr(self.keithley, smu_name)
        except Exception as e:
            print(f"获取SMU对象时出错: {str(e)}")
            self.smu = None
        self.interval = interval  # 数据更新间隔（秒）
        self.running = True
        self.time_data = []
        self.voltage_data = []
        self.current_data = []
        self.start_time = 0
        self.simulation_mode = simulation_mode
        
        # 动态电压值，可在测量过程中更新
        self.current_voltage = 1.0
        # 线程锁，防止电压更新和读取之间的竞争条件
        self.lock = QtCore.QMutex()

    def __del__(self):
        self.running = False
        self.wait()

    def stop(self):
        self.running = False
        
    def update_simulation_voltage(self, new_voltage):
        """更新模拟模式下的电压值"""
        self.lock.lock()
        self.current_voltage = new_voltage
        self.lock.unlock()
        
    def _run_simulation(self):
        """在模拟模式下生成随机数据"""
        import random
        import math
        
        self.time_data = []
        self.voltage_data = []
        self.current_data = []
        self.start_time = time.time()
        
        # 模拟参数
        frequency = 0.2  # 波动频率
        noise_level = 0.1  # 噪声水平
        baseline = 1e-6   # 基准电流值
        amplitude = 1e-6  # 波动幅度
        
        t = 0
        while self.running:
            # 获取当前电压值（使用锁防止竞态条件）
            self.lock.lock()
            simulated_voltage = self.current_voltage
            self.lock.unlock()
            
            # 电流随电压近似线性变化（欧姆定律模拟）的简化模型
            current_scale = abs(simulated_voltage) * 2e-6
            
            # 随机模拟电流值，包含一些周期性变化和噪声
            # 简单的正弦波加上随机噪声
            sine_component = amplitude * math.sin(2 * math.pi * frequency * t)
            noise = random.uniform(-noise_level, noise_level) * amplitude
            i = baseline + current_scale + sine_component + noise
            
            self.time_data.append(t)
            self.voltage_data.append(simulated_voltage)
            self.current_data.append(i)
            
            # 发送数据到主线程
            data = {
                "time": self.time_data,
                "voltage": self.voltage_data,
                "current": self.current_data,
                "last_value": (t, simulated_voltage, i)
            }
            self.data_sig.emit(data)
            
            # 等待指定间隔
            time.sleep(self.interval)
            t += self.interval

    def run(self):
        try:
            # 检查是否为模拟模式
            if self.simulation_mode:
                self._run_simulation()
                return
                
            # 以下是实际设备模式
            # 检查设备是否已连接
            if not hasattr(self.keithley, 'connected') or not self.keithley.connected:
                self.error_sig.emit(Exception("设备未连接"))
                return
                
            # 检查SMU是否有效
            if self.smu is None:
                self.error_sig.emit(Exception("无法获取SMU对象"))
                return
                
            # 设置更长的超时时间（如果可能）
            try:
                if hasattr(self.keithley, 'visa') and hasattr(self.keithley.visa, 'timeout'):
                    original_timeout = self.keithley.visa.timeout
                    # 设置为较长超时时间（30秒）
                    self.keithley.visa.timeout = 30000
            except Exception as e:
                print(f"设置超时时间出错: {str(e)}")
                
            self.time_data = []
            self.voltage_data = []
            self.current_data = []
            self.start_time = time.time()
            consecutive_errors = 0  # 连续错误计数
            max_consecutive_errors = 3  # 最大连续错误次数

            while self.running:
                try:
                    # 读取当前电压和电流
                    v = self.smu.measure.v()
                    i = self.smu.measure.i()
                    t = time.time() - self.start_time
                    
                    self.time_data.append(t)
                    self.voltage_data.append(v)
                    self.current_data.append(i)
                    
                    # 发送数据到主线程
                    data = {
                        "time": self.time_data,
                        "voltage": self.voltage_data,
                        "current": self.current_data,
                        "last_value": (t, v, i)
                    }
                    self.data_sig.emit(data)
                    
                    # 重置连续错误计数
                    consecutive_errors = 0
                    
                except pyvisa.VisaIOError as visa_error:
                    consecutive_errors += 1
                    print(f"VISA通信错误 ({consecutive_errors}/{max_consecutive_errors}): {str(visa_error)}")
                    
                    # 如果是超时错误，尝试恢复连接
                    if "VI_ERROR_TMO" in str(visa_error):
                        print("通信超时，尝试恢复...")
                        try:
                            # 短暂暂停
                            time.sleep(0.5)
                            # 尝试简单的读取操作来恢复连接
                            if hasattr(self.keithley, 'reset'):
                                print("正在重置设备连接...")
                                self.keithley.reset()
                                time.sleep(1)
                        except Exception as recover_error:
                            print(f"恢复连接尝试失败: {str(recover_error)}")
                    
                    # 如果连续错误次数超过阈值，退出循环
                    if consecutive_errors >= max_consecutive_errors:
                        self.error_sig.emit(Exception(f"多次测量失败，停止测量: {str(visa_error)}"))
                        break
                    
                    # 继续尝试下一次测量
                    continue
                        
                except Exception as e:
                    consecutive_errors += 1
                    print(f"测量过程中出错 ({consecutive_errors}/{max_consecutive_errors}): {str(e)}")
                    
                    # 如果连续错误次数超过阈值，退出循环
                    if consecutive_errors >= max_consecutive_errors:
                        self.error_sig.emit(e)
                        break
                    
                    # 继续尝试下一次测量
                    continue
                    
                # 等待指定间隔
                time.sleep(self.interval)
                
        except Exception as exc:
            self.error_sig.emit(exc)
        finally:
            # 恢复原始超时设置
            try:
                if hasattr(self.keithley, 'visa') and hasattr(self.keithley.visa, 'timeout') and 'original_timeout' in locals():
                    self.keithley.visa.timeout = original_timeout
            except Exception:
                pass
                
            self.finished_sig.emit()


def run():
    import sys
    import argparse
    from keithley2600 import log_to_screen

    parser = argparse.ArgumentParser()
    parser.add_argument(
        "-v", "--verbose", help="increase output verbosity", action="store_true"
    )
    args = parser.parse_args()
    if args.verbose:
        log_to_screen()

    app = QtWidgets.QApplication(sys.argv)

    try:
        keithley_gui = KeithleyGuiApp()
        keithley_gui.show()
        return app.exec()
    except Exception as e:
        print(f"启动应用程序时出错: {str(e)}")
        # 显示错误对话框
        error_dialog = QtWidgets.QMessageBox()
        error_dialog.setIcon(QtWidgets.QMessageBox.Critical)
        error_dialog.setWindowTitle("应用程序错误")
        error_dialog.setText("启动应用程序时发生错误")
        error_dialog.setDetailedText(f"{str(e)}")
        error_dialog.setStandardButtons(QtWidgets.QMessageBox.Ok)
        error_dialog.exec_()
        return 1


if __name__ == "__main__":

    run()
