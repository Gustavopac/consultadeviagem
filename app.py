from flask import Flask, render_template, request, redirect, url_for
import pandas as pd

app = Flask(__name__)

try:
    excel_data = pd.ExcelFile('comprovei_tracking.xlsx')
    if 'Dados Extraídos' in excel_data.sheet_names:
        dados_df = excel_data.parse('Dados Extraídos')
        print("Colunas encontradas:", dados_df.columns)
    else:
        print("Erro: A aba 'Dados Extraídos' não foi encontrada no arquivo.")
except FileNotFoundError:
    print("Erro: O arquivo comprovei_tracking.xlsx não foi encontrado.")

def buscar_dados(campo, valor):
    try:
        valor = str(valor).strip()

        if campo == "Número da nota" and "Número da nota" in dados_df.columns:
            resultado = dados_df[dados_df["Número da nota"].astype(str).str.strip() == valor]
        elif campo == "CTE" and "CTE" in dados_df.columns:
            resultado = dados_df[dados_df["CTE"].astype(str).str.strip() == valor]
        elif campo == "Manifesto + Parada" and "Manifesto + Parada" in dados_df.columns:
            resultado = dados_df[dados_df["Manifesto + Parada"].astype(str).str.strip() == valor]
        else:
            return None

        if not resultado.empty:
            link = resultado["Tracking"].values[0]
            return link
        else:
            return None
    except Exception as e:
        print("Erro durante a busca:", e)
        return None

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        campo = request.form["campo"]
        valor = request.form["valor"]

        link = buscar_dados(campo, valor)

        if link:
            return redirect(link)
        else:
            return render_template("index.html", error="Resultado não encontrado.")

    return render_template("index.html")

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000, debug=True)