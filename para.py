import subprocess
from kivy.uix.widget import Widget
from kivy.app import App
import paramiko
import os
                                        ###################################
                                        # ericsson OSS password reset     #
                                        # gui interface.                  #
                                        # you are free to modify this code#
                                        # EDOH ADEJO NEDOX                #
                                        # OSS Engineer @HUAWEI 2017       #
                                        ###################################                             
class OSSLook(Widget):

    def TerminalAccess(self):
        exit_code = 'exit'
        pwd = self.ids.password.text
        user = self.ids.username.text
        pwd2 = self.ids.password2.text
        exitcode = "exit"
        if pwd == pwd2:
            dssh = paramiko.SSHClient()
            dssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            dssh.connect('hostip', username='root', password='Adm@Inf1')
            stdin, stdout, stderr = dssh.exec_command("")
            stdin, stdout, stderr = dssh.exec_command('pwd')
            stdin, stdout, stderr = dssh.exec_command("cd /ericsson/sdee/bin; ./chg_user_password.sh")
            stdin.write('ldapadmin123')
            stdin.write('\n')
            stdin.flush()
            stdin.write(user)
            stdin.write('\n')
            stdin.flush()
            stdin.write(pwd)
            stdin.write('\n')
            stdin.flush()
            stdin.write(pwd)
            stdin.write('\n')
            stdin.flush()
            retValue = str(stdout.readlines()[9]).rstrip()
            print(retValue)
            stdin.write(exitcode)
            #print('INFO: Local user [' + user + '] Type [OSS_ONLY] Domain [zhomc]: Password changed.')
            if retValue == str('INFO: Local user [' + user + '] Type [OSS_ONLY] Domain [zhomc]: Password changed.'):
                print(retValue)
                self.ids.st.text = str("DONE")
            else:
                self.ids.st.text = str("try again")

            #self.ids.username.text = pwd
            #subprocess.Popen([r"C:\Program Files\Microsoft Office\Office15\EXCEL.exe"])


class OSSApp(App):
    def build(self):
        return OSSLook()


if __name__ == '__main__':
    OSSApp().run()
