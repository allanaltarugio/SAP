import win32com.client
import sys
import subprocess
import time
from tkinter import *
from tkinter import messagebox

class SapGui():
    def __init__(self):
        self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\sapshcut.exe"  # Use o sapshcut.exe
        subprocess.Popen([self.path, "-system=PRD", "-client=200", "-user=cs314073", "-pw=d@606060"])  # Inicie diretamente a conexão
        time.sleep(5)

        try:
            # Tente obter o objeto SAPGUI pelo caminho registrado
            self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
            if not type(self.SapGuiAuto) == win32com.client.CDispatch:
                raise Exception("Failed to get SAPGUI object")

            # Tente obter a aplicação SAP pela primeira vez
            self.application = self.SapGuiAuto.SAPGuiApp(0)
            if not type(self.application) == win32com.client.CDispatch:
                raise Exception("Failed to get application object")

            # Tente obter a conexão SAP pela primeira vez
            self.connection = self.application.OpenConnection("01 - Cosan / Raízen Energia – ECC (PRD)", True)
            if not type(self.connection) == win32com.client.CDispatch:
                raise Exception("Failed to get connection object")

            # Tente obter a sessão SAP pela primeira vez
            self.session = self.connection.Children(0)
            if not type(self.session) == win32com.client.CDispatch:
                raise Exception("Failed to get session object")

            # Maximize a janela
            self.session.findById("wnd[0]").maximize()
        except Exception as e:
            # print(f"Erro durante a inicialização: {e}")
            # messagebox.showerror("Erro", f"Erro durante a inicialização do SAP: {e}")
            sys.exit()

    def sapLogin(self):
        try:
            self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "200"
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "cs314073"
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "d@606060"
            self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "PT"
            self.session.findById("wnd[0]").sendVKey(0)
            messagebox.showinfo("Sucesso", "Login com sucesso!")

        except win32com.client.dynamic.pythoncom.com_error as e:
            print(f"Erro durante o login: {e}")
            messagebox.showerror("Erro", f"Erro durante o login: {e}")

if __name__ == "__main__":
    window = Tk()
    window.geometry("200x200")
    window.title("Starting SAP")  # Defina o título da janela
    window.configure(bg="#f0f0f0")  # Cor de fundo

    # Adicione uma marca d'água
    watermark_label = Label(window, text="Desenvolvedor: Allan Altarugio", font=("Arial", 9), fg="blue", bg="#f0f0f0")
    watermark_label.place(x=10, y=180)

    btn = Button(window, text="Login SAP", command=lambda: SapGui().sapLogin())
    btn.pack()

    window.mainloop()