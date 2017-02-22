from flask import Flask,request,redirect,url_for,send_from_directory
import crpi as cr
app = Flask(__name__)

OUTPUT_FOLDER = './output/'
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

@app.route('/extract/')
def extract():
	#with app.open_resource('./reports/SalesHistory.RPT') as f:
	try:
		report = request.args.get('report')	

		r = cr.connect('./reports/'+report+'.RPT')
		params = cr.get_params(r)

	except Exception as e:
		return str(e)

	else:
		val = ""
		for each in params:
			val += (",".join([cr.get_fieldname(each),cr.get_prompt_text(each), cr.get_valuetype(each)])) + ","
		return val
	

@app.route('/export/')
def export():	
	try:
		report = request.args.get('report')
		typ = request.args.get('type')
		values = cr.get_literals(request.args.get('values'))

		r = cr.connect('./reports/'+report+'.RPT')
		params = cr.get_params(r)	
		
		for each in params:
			each.ClearCurrentValueAndRange()
			each.AddCurrentValue(values.get(cr.get_fieldname(each),""))		
		filename = cr.set_exportoption(r,typ)
		r.Export(promptUser=False)

	except Exception as e:
		return str(e)
	
	else:		
		return redirect(url_for('uploaded_file',filename=filename))


@app.route('/outfiles/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['OUTPUT_FOLDER'], filename)

if __name__ == "__main__":
	app.run()