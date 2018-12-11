#A part of the AudioChart addon for NVDA
#Copyright (C) 2018 Tony Malykh
#This file is covered by the GNU General Public License.
#See the file COPYING.txt for more details.

import addonHandler
import api
import config
import controlTypes
import ctypes
import globalPluginHandler
import gui
from gui import guiHelper
import math
import NVDAHelper
from NVDAObjects.window import excel
import operator
import re 
import scriptHandler
from scriptHandler import script
import speech
import struct
import tones
import ui
import wx

addonHandler.initTranslation()

def initConfiguration():
    confspec = {
        "min_value" : "float( default=0)",
        "max_value" : "float( default=100)",
    }
    config.conf.spec["audiochart"] = confspec

initConfiguration()


beep_duration = 5
beep_volume = 50 # percent
mid_pitch = speech.IDT_BASE_FREQUENCY
beep_range = 4 # octaves
pitch_coefficient = 2**(beep_range / 2)
pitch_low = mid_pitch / pitch_coefficient
pitch_high = mid_pitch * pitch_coefficient

value_low = 0
value_high = 10

max_rows = 10000

def load_values():
    global value_low, value_high
    value_low = config.conf["audiochart"]["min_value"]
    value_high = config.conf["audiochart"]["max_value"]
load_values()

def value_to_pitch(value):
    z = (value - value_low) / (value_high - value_low)
    pitch = 2 ** (z * math.log(pitch_high/pitch_low, 2))
    pitch *= pitch_low
    pitch = max(pitch, 10)
    pitch = min(pitch, 30000)
    return pitch
    
class CalibrationDialog(wx.Dialog):
    def __init__(self, parent, values):
        # Translators: Title of calibration dialog
        title_string = _("AudioChart calibration")
        super(CalibrationDialog, self).__init__(parent, title=title_string)
        self.values = values
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        sHelper = gui.guiHelper.BoxSizerHelper(self, orientation=wx.VERTICAL)
        
        # Translators: Label for min value edit box
        minCtrlLabel = _("Min value")
        minCtrl = gui.guiHelper.LabeledControlHelper(self, minCtrlLabel, wx.TextCtrl)
        self.minEdit = minCtrl.control
        self.minEdit.Value = str(config.conf["audiochart"]["min_value"])

        # Translators: Label for max value edit box
        maxCtrlLabel = _("Max value")
        maxCtrl = gui.guiHelper.LabeledControlHelper(self, maxCtrlLabel, wx.TextCtrl)
        self.maxEdit = maxCtrl.control
        self.maxEdit.Value = str(config.conf["audiochart"]["max_value"])

        bHelper = sHelper.addItem(guiHelper.ButtonHelper(orientation=wx.HORIZONTAL))
        # Translators: Calibrate button text
        calibrateLabel = _("&Calibrate")
        self.calibrateButton = bHelper.addButton(self, label=calibrateLabel)
        self.Bind(wx.EVT_BUTTON, self.onCalibrate, self.calibrateButton)
        
        sHelper.addDialogDismissButtons(self.CreateButtonSizer(wx.OK|wx.CANCEL))
        self.Bind(wx.EVT_BUTTON, self.onOk, id=wx.ID_OK)
        mainSizer.Add(sHelper.sizer, border=gui.guiHelper.BORDER_FOR_DIALOGS, flag=wx.ALL)
        self.Sizer = mainSizer
        mainSizer.Fit(self)
        #self.calibrateButton.SetFocus()
        self.CentreOnScreen()

    def onOk(self, evt):
        if not self.validate():
            return
        self.saveSettings()
        self.Destroy()
        playAsync(self.values)
        
    def onCalibrate(self, evt):
        self.minEdit.Value = str(min(self.values))
        self.maxEdit.Value = str(max(self.values))
        
    def saveSettings(self):
        config.conf["audiochart"]["min_value"] = float(self.minEdit.Value)
        config.conf["audiochart"]["max_value"] = float(self.maxEdit.Value)
        load_values()
        
    def validate(self):
        try:
            min_value = self.minEdit.Value
            assert(not math.isnan(min_value))
            assert(not math.isinf(min_value))
        except:
            self.minEdit.SetFocus()
            # Translators: message when min value is invalid
            ui.message(_("Invalid min value."))
            return False
        try:
            max_value = self.maxEdit.Value
            assert(not math.isnan(max_value))
            assert(not math.isinf(max_value))
        except:
            self.maxEdit.SetFocus()
            # Translators: message when max value is invalid
            ui.message(_("Invalid max value."))
            return False
        if not (min_value < max_value):
            # Translators: message when min value is greater than max value
            ui.message(_("Min value should be less than max value."))
            return False
        return True

            
def playAsync(values):
    wx.CallAfter(play, values)

def play(values):
    pitches = list(map(value_to_pitch, values))
    n = len(pitches)
    beepBufSizes = [NVDAHelper.generateBeep(None, pitches[i], beep_duration, beep_volume, beep_volume) for i in range(n)]
    bufSize = sum(beepBufSizes)
    buf = ctypes.create_string_buffer(bufSize)
    bufPtr = 0
    for pitch in pitches:
        bufPtr += NVDAHelper.generateBeep(
            ctypes.cast(ctypes.byref(buf, bufPtr), ctypes.POINTER(ctypes.c_char)), 
            pitch, beep_duration, beep_volume, beep_volume)
    tones.player.stop()
    speech.cancelSpeech()
    tones.player.feed(buf.raw)


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    scriptCategory = _("AudioChart")
    
    @script(description='Plot audio chart.', gestures=['kb(desktop):nvda+a', 
        'kb(laptop):nvda+control+shift+a'])
    def script_audioChart(self, gesture):
        count=scriptHandler.getLastScriptRepeatCount()
        if count >= 2:
            return
        values = self.collectValues()
        if values is None:
            return
        if count == 0:
            playAsync(values)
        else:
            tones.player.stop()
            self.showCalibrationDialog(values)
        
    def collectValues(self):
        focus = api.getFocusObject()
        values = []
        if isinstance(focus, excel.ExcelSelection):
            colspan = focus._get_colSpan()
            rowspan = focus._get_rowSpan()
            if (colspan != 1) and (rowspan != 1):
                # Translators: Message when more than one column and more than 1 row is selected
                ui.message(_("Please select only a single column or a single row."))
                return None
            excelValues = focus.excelRangeObject.Value()
            for evTuple in excelValues:
                for ev in evTuple:
                    try:
                        values.append(float(ev))
                    except:
                        continue
            if len(values) == 0:
                # Translators: message when no numeric cells found within selected range
                ui.message(_("No numeric values found within the selection."))
                return None
        elif isinstance(focus, excel.ExcelCell):
            excelValues = focus.excelCellObject.Range("A1","A%d" % max_rows).Value()
            for evTuple in excelValues:
                try:
                    ev = evTuple[0]
                    values.append(float(ev))
                except:
                    break
            if len(values) == 0:
                # Translators: message when currently selected cell doesn't contain a numeric value
                ui.message(_("Please select a numeric value - the beginning of time series."))
                return None
        else:
            # Translators: message when called not in Excel worksheet
            ui.message(_("Audio chart currently only works inMicrosoft Excel."))
            return None
        return values
        
    def collectAndPlay(self):
        values = self.collectValues()
        play(values)
        
    def popupCalibrationDialog(self, values):
        gui.mainFrame.prePopup()
        d = CalibrationDialog(gui.mainFrame, values)
        d.Show()
        gui.mainFrame.postPopup()
        
    def showCalibrationDialog(self, values):
        wx.CallAfter(self.popupCalibrationDialog, values)
