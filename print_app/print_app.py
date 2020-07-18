'''
    Primary AutoPrint V 1.0
     - Quick turnaround auto-print program
'''
import imaplib
import email
import os
import tempfile
import shutil
import subprocess
import time
import winreg
import glob

def main():
    print('Welcome')
    print('\n Enter Log-in Email')

    user = input('     E-mail:')
    password = input('     Password:')
    imap_url = 'imap.gmail.com'

    ticker = 0

    while ticker < 3600:
        imap_url = 'imap.gmail.com'
        box = imaplib.IMAP4_SSL(imap_url)
        dir = tempfile.mkdtemp()
        AcroRD32Path = winreg.QueryValue(winreg.HKEY_CLASSES_ROOT, 'Software\\Adobe\\Acrobat\Exe')
        acroread = AcroRD32Path

        try:
            box = imaplib.IMAP4_SSL(imap_url)
            box.login(user, password)
        except:
            imaplib.error: b'[AUTHENTICATIONFAILED] Invalid credentials (Failure)'
            if ticker == 0:
                print("Login Failed")
                return main()
            else:
                print('Ooops Something went wrong \n Please Sign In again.')
                return main()
            break

        if ticker == 0:
            print( '      \nAuto-Printing has Started')
            print( '      \nAuto-Print will STOP in 10 hours')
            print( '      \nTo STOP at any time press CTRL+C')

        box.select('INBOX')
        tmp, data = box.search(None, '(FROM "Email_Server")')

        ticker = ticker + 1

        if data == [b'']:
            time.sleep(10)

        for num in data[0].split():
            tmp, data = box.fetch(num, '(RFC822)')
            raw = email.message_from_bytes(data[0][1])

            for part in raw.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue
                fileName = part.get_filename()
                if bool(fileName):
                    filePath = os.path.join(dir, fileName)
                    with open(filePath, 'wb') as f:
                        f.write(part.get_payload(decode=True))
            box.store(num, '+X-GM-LABELS', '\\Trash')
        box.expunge()


        # QUE TO PRINTING
        que = glob.glob(os.path.join(dir, '*.pdf' ))
        for report in que:
            cmd = '{0} /h /P "{1}" ""'.format(acroread, report)
            proc = subprocess.Popen(cmd)

            # Wait for PDF to be sent to printer
            print('Report Printed')
            time.sleep(10)
            os.remove(report)

        if ticker = 3600:
            print('Session Expired')
            shutil.rmtree(dir)
            sys.exit()

if __name__=='__main__':
    main()