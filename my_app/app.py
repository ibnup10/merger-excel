import os
import pandas as pd
from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
ALLOWED_EXTENSIONS = {'xlsx'}

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def combine_excel_sheets(folder_path, output_filename):
    """
    Menggabungkan semua sheet dari berbagai file Excel dalam satu folder menjadi satu file Excel baru,
    menjaga format angka pada kolom "NOMOR AJU" tetap utuh, termasuk angka nol di awal,
    dan menghapus kolom "Source.Name" jika ada.
    """
    data_dict = {}
    all_files = os.listdir(folder_path)
    excel_files = [f for f in all_files if f.endswith('.xlsx')]

    if not excel_files:
        return None, "Tidak ada file Excel ditemukan di folder yang diunggah."

    for file in excel_files:
        file_path = os.path.join(folder_path, file)
        try:
            excel_data = pd.read_excel(file_path, sheet_name=None, dtype={'NOMOR AJU': str})
            for sheet_name, df in excel_data.items():
                if sheet_name not in data_dict:
                    data_dict[sheet_name] = []
                if 'Source.Name' in df.columns:
                    df.drop(columns=['Source.Name'], inplace=True)
                data_dict[sheet_name].append(df)
        except Exception as e:
            return None, f"Terjadi error saat membaca file {file}: {str(e)}"
        finally:
            os.remove(file_path) # Bersihkan file setelah diproses

    if not data_dict:
        return None, "Tidak ada data yang berhasil digabungkan."

    output_file_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
    try:
        with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
            for sheet_name, df_list in data_dict.items():
                combined_df = pd.concat(df_list, ignore_index=True)
                combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
        return output_file_path, None
    except Exception as e:
        return None, f"Terjadi error saat menulis file hasil: {str(e)}"

@app.route('/', methods=['GET'])
def upload_form():
    return render_template('upload.html')

@app.route('/combine', methods=['POST'])
def combine_files():
    if 'excel_files' not in request.files:
        return render_template('upload.html', error='Tidak ada folder yang diunggah.')

    files = request.files.getlist('excel_files')
    output_filename = request.form.get('output_filename', 'Combined_Data.xlsx')
    temp_folder = os.path.join(app.config['UPLOAD_FOLDER'], 'temp_' + str(os.urandom(8).hex()))
    os.makedirs(temp_folder)
    uploaded_files = []
    error = None

    try:
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file_path = os.path.join(temp_folder, filename)
                file.save(file_path)
                uploaded_files.append(file_path)
            elif file:
                error = 'Hanya file dengan ekstensi .xlsx yang diperbolehkan.'
                break # Stop jika ada file yang tidak sesuai

        if not error and uploaded_files:
            output_file, combine_error = combine_excel_sheets(temp_folder, output_filename)
            if output_file:
                return send_file(output_file, as_attachment=True, download_name=output_filename)
            else:
                error = combine_error
        elif not files:
            error = 'Tidak ada file yang dipilih.'

    except Exception as e:
        error = f"Terjadi error: {str(e)}"
    finally:
        # Bersihkan folder temporary
        if os.path.exists(temp_folder):
            for item in os.listdir(temp_folder):
                item_path = os.path.join(temp_folder, item)
                if os.path.isfile(item_path):
                    os.remove(item_path)
            os.rmdir(temp_folder)

    return render_template('upload.html', error=error)

if __name__ == '__main__':
    app.run(debug=True)