import win32com.client as win32

class OutlookEmailSender:
    def send_email(self, body):
        try:
            outlook = win32.gencache.EnsureDispatch("Outlook.Application")
            email = outlook.CreateItem(8)  # 8 represents the MailItem type (email)

            email.subject = "Link Status Update"
            email.Body = "\n".join(body)
            email.To = "ajeesh.kumang@lloydsbanking.com"
            email.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/string/{68026386-0060-0060-C060-000609000846}/sip_labels/@x0080001F", "MSIP_Label_7bc792f8-6d75-423a-9981-629281829092_Enabled=true;")

            email.send()
            print("Email sent successfully.")
        except Exception as e:
            print(f"An error occurred while sending the email: {e}")
