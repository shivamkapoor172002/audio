from flask import Flask, render_template, request

app = Flask(__name__, static_folder='static', template_folder='templates')

@app.route('/')
def home():
    return render_template('home.html')

@app.route('/about')
def about():
    return render_template('about.html')

@app.route('/services')
def services():
    return render_template('services.html')

@app.route('/contact', methods=['GET', 'POST'])
def contact():
    form = {"name": "", "email": "", "subject": "", "message": ""}
    errors = []
    success = False

    if request.method == 'POST':
        form["name"] = request.form.get('name', '').strip()
        form["email"] = request.form.get('email', '').strip()
        form["subject"] = request.form.get('subject', '').strip()
        form["message"] = request.form.get('message', '').strip()

        if not form["name"]:
            errors.append("Name is required.")
        if not form["email"] or "@" not in form["email"]:
            errors.append("A valid email is required.")
        if not form["subject"]:
            errors.append("Subject is required.")
        if not form["message"]:
            errors.append("Message is required.")

        if not errors:
            success = True
            form = {"name": "", "email": "", "subject": "", "message": ""}

    return render_template('contact.html', form=form, errors=errors, success=success)

if __name__ == '__main__':
    app.run(debug=True, port=5000)

# Export for Vercel
app = app
