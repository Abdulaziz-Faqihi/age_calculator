import wx
from datetime import datetime
import ctypes
import os

# Load the NVDA controller client DLL
nvda = ctypes.WinDLL(os.path.join(os.path.dirname(__file__), 'nvdaControllerClient.dll'))

# Define the function from the DLL
nvda.nvdaController_speakText.argtypes = [ctypes.c_wchar_p]
nvda.nvdaController_speakText.restype = ctypes.c_int

class AgeCalculatorFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title='Age Calculator', size=(400, 300))
        
        panel = wx.Panel(self)
        
        sizer = wx.BoxSizer(wx.VERTICAL)
        
        self.dob_label = wx.StaticText(panel, label='Enter your Date of Birth:')
        sizer.Add(self.dob_label, 0, wx.ALL | wx.CENTER, 5)
        
        self.day_spin = wx.SpinCtrl(panel, value='1', min=1, max=31)
        sizer.Add(self.day_spin, 0, wx.ALL | wx.CENTER, 5)
        
        self.month_spin = wx.SpinCtrl(panel, value='1', min=1, max=12)
        sizer.Add(self.month_spin, 0, wx.ALL | wx.CENTER, 5)
        
        self.year_spin = wx.SpinCtrl(panel, value='2000', min=1900, max=datetime.today().year)
        sizer.Add(self.year_spin, 0, wx.ALL | wx.CENTER, 5)
        
        self.calculate_button = wx.Button(panel, label='Calculate Age')
        self.calculate_button.Bind(wx.EVT_BUTTON, self.on_calculate)
        sizer.Add(self.calculate_button, 0, wx.ALL | wx.CENTER, 5)
        
        panel.SetSizer(sizer)
        
    def on_calculate(self, event):
        try:
            day = self.day_spin.GetValue()
            month = self.month_spin.GetValue()
            year = self.year_spin.GetValue()
            dob = datetime(year, month, day)
            age_details = self.calculate_age(dob)
            result_text = (
                f'Your age is: {age_details["years"]} years, '
                f'{age_details["months"]} months, '
                f'{age_details["days"]} days, '
                f'{age_details["hours"]} hours, and '
                f'{age_details["minutes"]} minutes'
            )
            nvda.nvdaController_speakText(result_text)
        except ValueError:
            nvda.nvdaController_speakText('Invalid Date! Please enter valid numbers for day, month, and year.')
    
    def calculate_age(self, dob):
        now = datetime.now()
        delta = now - dob
        years = delta.days // 365
        months = (delta.days % 365) // 30
        days = (delta.days % 365) % 30
        hours = delta.seconds // 3600
        minutes = (delta.seconds % 3600) // 60
        return {
            "years": years,
            "months": months,
            "days": days,
            "hours": hours,
            "minutes": minutes
        }

if __name__ == '__main__':
    app = wx.App(False)
    frame = AgeCalculatorFrame()
    frame.Show()
    app.MainLoop()
