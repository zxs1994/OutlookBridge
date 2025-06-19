Set outlook = CreateObject("Outlook.Application")
Set mail = outlook.CreateItem(0)
mail.Subject = "VBScript Test"
mail.Body = "This is a VBScript test."
mail.Display