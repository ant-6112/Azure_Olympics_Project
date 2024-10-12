import win32com.client as win32

class EmailHandler:
    def __init__(self):
        self.outlook = win32.Dispatch('outlook.application')

    def send_email(self, recipient, body):
        try:
            mail = self.outlook.CreateItem(0)
            mail.To = recipient
            mail.Subject = "Risk Report Notification"
            mail.Body = body
            mail.Send()
            return True
        except Exception as e:
            print(f"Failed to send email: {str(e)}")
            return False
