import os
import win32com
import win32com.client


def remove_headers_footers(fname):
    fname = os.path.abspath(fname)
    wapp = win32com.client.gencache.EnsureDispatch('Word.Application')
    wconst = win32com.client.constants
    wapp.Visible = 0
    doc = wapp.Documents.Open(fname)
    for section in wapp.ActiveDocument.Sections:
        section.Headers(wconst.wdHeaderFooterPrimary).Range.Text = ''
        section.Footers(wconst.wdHeaderFooterPrimary).Range.Text = ''
    doc.Save()
    doc.Close()
    wapp.Quit()


def batch_remove(path, suffix=('doc', 'docx')):
    for root, dirs, fname in os.walk(path):
        for name in fname:
            if not name.startswith('~$') and name.endswith(suffix):
                print('Processing %s', os.path.join(root, name))
                remove_headers_footers(os.path.join(root, name))


batch_remove('.')
