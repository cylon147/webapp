from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory
import pandas as pd

app = Flask(__name__)
app.secret_key = 'supersecretkey'

# Load the Excel file
excel_file = 'data/guests.xlsx'
df = pd.read_excel(excel_file)

# Strip leading/trailing spaces from column names
df.columns = [col.strip() for col in df.columns]

@app.route('/')
def home():
    return render_template('start.html')

@app.route('/start', methods=['POST'])
def start():
    sr_no = request.form['sr_no']
    try:
        sr_no = int(sr_no)  # Ensure the input is an integer
        entry_index = df[df['Sr. No.'] == sr_no].index[0]
        return redirect(url_for('form', entry_index=entry_index))
    except IndexError:
        flash("Sr. No. not found.")
        return redirect(url_for('home'))
    except KeyError as e:
        flash(f"KeyError: {e}. Available columns: {df.columns}")
        return redirect(url_for('home'))
    except ValueError:
        flash("Please enter a valid number for Sr. No.")
        return redirect(url_for('home'))
    except Exception as e:
        flash(f"Unexpected error: {e}")
        return redirect(url_for('home'))

@app.route('/form/<int:entry_index>', methods=['GET', 'POST'])
def form(entry_index):
    if request.method == 'POST':
        try:
            for col in df.columns:
                if col not in ['Sr. No.', 'First Name', 'Middle Name', 'Last Name', 'Primary Phone', 'Mobile Phone', 'Home Street', 'Home City']:
                    df.at[entry_index, col] = request.form.get(col, df.at[entry_index, col])
            df.to_excel(excel_file, index=False)
            flash("Data saved successfully.")
        except Exception as e:
            flash(f"Error saving data: {e}")

        next_index = entry_index + 1
        if next_index >= len(df):
            return redirect(url_for('home'))
        else:
            return redirect(url_for('form', entry_index=next_index))

    # Replace NaN with empty string in the person data
    person = df.iloc[entry_index].fillna('')
    return render_template('form.html', entry_index=entry_index, person=person, columns=df.columns)

@app.route('/back/<int:entry_index>', methods=['POST'])
def back(entry_index):
    prev_index = entry_index - 1
    if prev_index < 0:
        return redirect(url_for('home'))
    return redirect(url_for('form', entry_index=prev_index))

@app.route('/goto', methods=['POST'])
def goto():
    sr_no = request.form['sr_no']
    try:
        sr_no = int(sr_no)  # Ensure the input is an integer
        entry_index = df[df['Sr. No.'] == sr_no].index[0]
        return redirect(url_for('form', entry_index=entry_index))
    except IndexError:
        flash("Sr. No. not found.")
        return redirect(url_for('home'))
    except KeyError as e:
        flash(f"KeyError: {e}. Available columns: {df.columns}")
        return redirect(url_for('home'))
    except ValueError:
        flash("Please enter a valid number for Sr. No.")
        return redirect(url_for('home'))
    except Exception as e:
        flash(f"Unexpected error: {e}")
        return redirect(url_for('home'))

@app.route('/download')
def download_file():
    return send_from_directory(directory='data', path='guests.xlsx', as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
