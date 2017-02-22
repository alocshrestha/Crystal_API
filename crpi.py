#dcr stuff here..
import sys, ast, uuid
import win32com
from win32com.client import Dispatch

def connect(rpt):	
	rep = openreport(rpt)
	tbl = rep.Database.Tables.Item(1)
	prop = tbl.ConnectionProperties('Password')
	prop.Value = ""
	prop = tbl.ConnectionProperties('Server')
	prop.Value = ''
	return rep

def openreport(name):
	app = Dispatch('CrystalRunTime.Application')
	return app.OpenReport(name)

def get_params(self):
	return self.ParameterFields

#do not change the constants; ref:CRFieldValueType
def get_valuetype(self):
	vtype = "none"
	if self.valuetype == 12:
		vtype = "string"
	elif self.valuetype == 9:
		vtype == "bool"
	elif self.valuetype == 8:
		vtype == "curr"
	elif self.valuetype == 10:
		vtype = "date"
	elif self.valuetype == 16:
		vtype = "datetime"
	elif self.valuetype == 7:
		vtype = "number"
	elif self.valuetype == 11:
		vtype = "time"
	return vtype

def get_fieldname(self):
	return self.parameterfieldname

def get_prompt_text(self):
	return self.prompt

def get_literals(self):
	return ast.literal_eval(self)

def gen_filekey():
	x = uuid.uuid4()
	return str(x)

#do not change the constants; ref:CRExportFormatType
def set_exportoption(r, typ):
	exp = r.ExportOptions
	key = gen_filekey()
	filename = ""
	if typ in "csv":		
		exp.DestinationType = 1
		exp.DiskFileName = './output/'+key+'.csv'
		exp.FormatType = 7
		filename = key + '.csv'
	elif typ in "pdf":
		exp.DestinationType = 1
		exp.FormatType = 31
		exp.DiskFileName = './output/'+key+'.pdf'
		filename = key + '.pdf'
	elif typ in "excel":
		exp.DestinationType = 1
		exp.FormatType = 36
		exp.DiskFileName = './output/'+key+'.xls'
		filename = key + '.xls'
	return filename











