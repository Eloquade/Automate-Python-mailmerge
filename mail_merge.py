import os
import win32com.client as win32 # pip install pywin32

working_directory = os.getcwd()
source_name = 'Data Source.xlsx'
destination_folder = os.path.join(working_directory, 'Destination')

"""
Create a Word application instance
"""
wordApp = win32.Dispatch('Word.Application')
wordApp.Visible = True

"""
Open Word Template + Open Data Source
"""
sourceDoc = wordApp.Documents.Open(os.path.join(working_directory, 'Word Template.docx'))
mail_merge = sourceDoc.MailMerge
mail_merge.OpenDataSource(
    Name:=os.path.join(working_directory, source_name),
    sqlstatement:="SELECT * FROM [Data Source$]"
)

record_count = mail_merge.DataSource.RecordCount

"""
Perform Mail Merge
"""
for i in range(1, record_count + 1):
    mail_merge.DataSource.ActiveRecord = i
    mail_merge.DataSource.FirstRecord = i
    mail_merge.DataSource.LastRecord = i

    mail_merge.Destination = 0
    mail_merge.Execute(False)

    # get record value
    base_name = mail_merge.DataSource.DataFields('Name'.replace(' ', '_')).Value

    targetDoc = wordApp.ActiveDocument

    """
    Save Files in Word Doc and PDF
    """
    targetDoc.SaveAs2(os.path.join(destination_folder, base_name + '.docx'))
    targetDoc.ExportAsFixedFormat(os.path.join(destination_folder, base_name), exportformat:=17)
    
    """
    Close target file
    """
    targetDoc.Close(False)
    targetDoc = None
    
sourceDoc.MailMerge.MainDocumentType = -1