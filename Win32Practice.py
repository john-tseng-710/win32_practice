import win32api, win32gui, win32con
import win32com.client
import pyautogui


class Win32Practice(object):

    wait_time = 2000

    def __init__(self, calc_input, notepad_input, word_input, excel_input):
        self.calc(calc_input)
        self.notepad(notepad_input, saveas=r'C:\Users\user1\Desktop\hello.txt')
        self.word(word_input, saveas=r'C:\Users\user1\Desktop\hello.docx')
        self.excel('A', excel_input, saveas=r'C:\Users\user1\Desktop\hello.xlsx')

    @classmethod
    def calc(cls, input):
        wscript = win32com.client.Dispatch('WScript.Shell')
        wscript.run('calc')
        win32api.Sleep(500)
        win = win32gui.FindWindow(None, '小算盤')
        win32gui.SetForegroundWindow(win)
        win32api.Sleep(500)
        for i in input:
            if i != ' ' or i != '=':
                pyautogui.press(i)
        pyautogui.press('=')
        win32api.Sleep(cls.wait_time)
        win32gui.PostMessage(win, win32con.WM_CLOSE, 0, 0)


    @classmethod
    def notepad(cls, input, saveas=None):
        wscript = win32com.client.Dispatch('WScript.Shell')
        wscript.run('notepad')
        win32api.Sleep(500)
        win = win32gui.FindWindow(None, '未命名 - 記事本')
        win32gui.SetForegroundWindow(win)
        win32api.Sleep(500)
        pyautogui.hotkey('ctrl', 'space')
        for i in input:
            pyautogui.press(i)
        win32api.Sleep(500)
        win32gui.PostMessage(win, win32con.WM_CLOSE, 0, 0)
        if saveas:
            pyautogui.hotkey('alt', 's')
            win32api.Sleep(500)
            pyautogui.write(str(saveas))
            pyautogui.hotkey('alt', 's')

    @ classmethod
    def word(cls, input, saveas=None):
        app = win32com.client.Dispatch('Word.Application')
        app.Visible = 1
        doc = app.Documents.Add()
        app.Selection.Text = input
        doc.SaveAs(saveas)
        doc.Close()
        app.Quit()

    @classmethod
    def excel(cls, col, input, saveas=None):
        app = win32com.client.Dispatch('Excel.Application')
        app.Visible = 1
        wb = app.Workbooks.Add()
        ws = wb.Worksheets('工作表1')
        for r in range(len(input)):
            pos = col + str(r+1)
            ws.Range(pos).value = input[r]
        wb.SaveAs(saveas)
        wb.Close()
        app.Quit()

# if __name__ == '__main__':
#     Practice('1+1', 'Hi, notepad.', 'Hi, word.', 'Hi, excel.')