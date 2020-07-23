from flask import Flask, render_template, request
from excel import by_index

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def main():
    if request.method == 'POST':
        if 'sheet' in request.files:
            file = request.files['sheet']
            fname = 'uploaded_file.xlsx'
            file.save(fname)
            fields = request.form.get('fields', '').split('\n')
            res = by_index(fname, request.form.get('list'), request.form.get('x_start'), 
                request.form.get('x_stop'), int(request.form.get('y_start')), int(request.form.get('y_stop')), 
                fields
            )
            return render_template('root.html', res=res)
    return render_template('root.html')

if __name__ == '__main__':
    app.run(debug=True)