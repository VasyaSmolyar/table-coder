from flask import Flask, render_template, request
from excel import by_name, by_index, tree_by_name

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def main():
    args = ['' for i in range(7)]
    if request.method == 'POST':
        if 'sheet' in request.files:
            file = request.files['sheet']
            fname = 'uploaded_file.xlsx'
            file.save(fname)
            fields = request.form.get('fields', '').split('\n')
            res = by_name(fname, request.form.get('list'), request.form.get('x_field'), request.form.get('x_start'), 
                request.form.get('x_stop'), int(request.form.get('y_start')), int(request.form.get('y_stop')), 
                fields
            )
            args = [request.form.get('list'), request.form.get('x_start'), 
                request.form.get('x_stop'), request.form.get('y_start'), request.form.get('y_stop'), 
                request.form.get('fields', ''), request.form.get('x_field', '')]
            return render_template('root.html', res=res, args=args, name="По заголовку")
    return render_template('root.html', args=args, name="По заголовку")

@app.route('/index', methods=['GET', 'POST'])
def index():
    args = ['' for i in range(7)]
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
            args = [request.form.get('list'), request.form.get('x_start'), 
                request.form.get('x_stop'), request.form.get('y_start'), request.form.get('y_stop'), request.form.get('fields', '')]
            return render_template('root.html', res=res, args=args, name="По индексу")
    return render_template('root.html', args=args, name="По индексу")

@app.route('/tree/name', methods=['GET', 'POST'])
def tree_name():
    args = ['' for i in range(7)]
    if request.method == 'POST':
        if 'sheet' in request.files:
            file = request.files['sheet']
            fname = 'uploaded_file.xlsx'
            file.save(fname)
            fields = request.form.get('fields', '').split('\n')
            res = tree_by_name(fname, request.form.get('list'), request.form.get('x_start'), 
                request.form.get('x_stop'), int(request.form.get('y_start')), int(request.form.get('y_stop')), 
                fields
            )
            args = [request.form.get('list'), request.form.get('x_start'), 
                request.form.get('x_stop'), request.form.get('y_start'), request.form.get('y_stop'), request.form.get('fields', '')]
            return render_template('root.html', res=res, args=args, name="Дерево имён")
    return render_template('root.html', args=args, name="Дерево имён")

if __name__ == '__main__':
    app.run(debug=True)