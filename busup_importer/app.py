import os
import io
import re
import unicodedata
import pandas as pd
from flask import Flask, render_template, request, send_file, flash, redirect, url_for

app = Flask(__name__)
app.secret_key = "busup_secret_key_2026"

# Permite encontrar a base tanto na subpasta 'bases' quanto na raiz do projeto
BASE_ROUTES_PATHS = [
    os.path.join('bases', 'base_route_stops.xlsx'),
    'base_route_stops.xlsx'
]

# Dicionário de Mapeamento dos Sites (Amazon e Mercado Livre) para company_id
COMPANY_SITE_MAP = {
    # Amazon
    'GRU8': 794,
    'GRU9': 1022,
    'GRU5': 793,

    # Mercado Livre
    '818': 818,
    '28 - Mercado Livre - Uberaba - MG 12': 818,
    'UBERABA - MG 12': 818,

    '829': 829,
    '28 - Mercado Livre - Barueri - SP 07': 829,
    'BARUERI - SP 07': 829,

    '837': 837,
    '28 - Mercado Livre - Betim - MG 02': 837,
    'BETIM - MG 02': 837,

    '1110': 1110,
    '28 - Mercado Livre - Embu das Artes - RC 03': 1110,
    'EMBU DAS ARTES - RC 03': 1110,

    '1172': 1172,
    '1242': 1242,
    '28 - Mercado Livre - Cotia - SP 25/26': 1172,
    'COTIA - SP 25/26': 1172,

    '1174': 1174,
    '28 - Mercado Livre - Flores da Cunha - SR S4': 1174,
    'FLORES DA CUNHA - SR S4': 1174,

    '1185': 1185,
    '28 - Mercado Livre - Passo Fundo - SR S5': 1185,
    'PASSO FUNDO - SR S5': 1185,

    '1186': 1186,
    '28 - Mercado Livre - Pelotas - SR S2': 1186,
    'PELOTAS - SR S2': 1186,

    '1187': 1187,
    '28 - Mercado Livre - Santa Maria - SR S3': 1187,
    'SANTA MARIA - SR S3': 1187,

    '1188': 1188,
    '28 - Mercado Livre - Contagem - SM G1': 1188,
    'CONTAGEM - SM G1': 1188,

    '1189': 1189,
    '28 - Mercado Livre - Contagem - XM G1': 1189,
    'CONTAGEM - XM G1': 1189,

    '1190': 1190,
    '28 - Mercado Livre - Contagem - SM G15': 1190,
    'CONTAGEM - SM G15': 1190,

    '1191': 1191,
    '28 - Mercado Livre - Vespasiano - SM G08': 1191,
    'VESPASIANO - SM G08': 1191,

    '1264': 1264,
    '28 - Mercado Livre - Contagem - SM G18': 1264,
    'CONTAGEM - SM G18': 1264,

    '1295': 1295,
    '28 - Mercado Livre - Cotia - SP 56': 1295,
    'COTIA - SP 56': 1295,

    '1300': 1300,
    '28 - Mercado Livre - Barueri - SP 58': 1300,
    'BARUERI - SP 58': 1300,

    '1332': 1332,
    '28 - Mercado Livre - Barueri - SP 60': 1332,
    'BARUERI - SP 60': 1332
}

MELI_COMPANY_IDS = {
    818, 829, 837, 1110, 1172, 1242, 1174, 1185, 1186, 1187, 
    1188, 1189, 1190, 1191, 1264, 1295, 1300, 1332
}

TEMPLATE_COLUMNS = [
    'user_id', 'name', 'last_name', 'address', 'address_latitude', 
    'address_longitude', 'neighborhood', 'postal_code', 'city', 'state', 
    'email', 'phone_number', 'password', 'company_id', 'shifts', 
    'cost_center_id', 'document', 'number_worker', 'position', 'start_date', 
    'route_going_id', 'route_return_id', 'rate_going_id', 'rate_return_id', 
    'stop_origin_going_id', 'stop_destination_going_id', 
    'stop_origin_return_id', 'stop_destination_return_id', 'user_plan_id'
]

def normalizar_texto(texto):
    """Remove acentos, espaços extras, trata números e converte para maiúsculo."""
    if pd.isna(texto) or texto is None:
        return ''
    if isinstance(texto, float) and texto.is_integer():
        texto = int(texto)
    texto_str = str(texto).upper()
    nfkd = unicodedata.normalize('NFD', texto_str)
    sem_acento = "".join([c for c in nfkd if not unicodedata.combining(c)])
    return re.sub(r'[^A-Z0-9]', '', sem_acento)

def obter_caminho_base():
    for caminho in BASE_ROUTES_PATHS:
        if os.path.exists(caminho):
            return caminho
    return None

def carregar_base_rotas():
    caminho = obter_caminho_base()
    if caminho:
        try:
            df = pd.read_excel(caminho)
            df.columns = [str(c).strip() for c in df.columns]
            for col in df.select_dtypes(include=['object']).columns:
                df[col] = df[col].fillna('').astype(str).str.strip()

            if 'Route name' in df.columns:
                df['Route_name_norm'] = df['Route name'].apply(normalizar_texto)
            if 'Route code' in df.columns:
                df['Route_code_norm'] = df['Route code'].apply(normalizar_texto)
            if 'Stop number' in df.columns:
                df['Stop_number_num'] = pd.to_numeric(df['Stop number'], errors='coerce')

            return df
        except Exception as e:
            print("Erro ao carregar base de rotas:", e)
            return None
    return None

def encontrar_coluna(df, nomes_possiveis):
    cols_originais = df.columns
    cols_map = {normalizar_texto(c): c for c in cols_originais}
    for nome in nomes_possiveis:
        chave_busca = normalizar_texto(nome)
        if chave_busca in cols_map:
            return cols_map[chave_busca]
    return None

def obter_company_id(site_nome):
    if pd.isna(site_nome) or not site_nome:
        return ''
    site_norm = normalizar_texto(site_nome)
    if not site_norm:
        return ''

    chaves_ordenadas = sorted(COMPANY_SITE_MAP.keys(), key=len, reverse=True)
    for site_key in chaves_ordenadas:
        key_norm = normalizar_texto(site_key)
        if key_norm and key_norm in site_norm:
            return COMPANY_SITE_MAP[site_key]
    return ''

@app.route('/')
def index():
    base_existe = obter_caminho_base() is not None
    return render_template('index.html', base_existe=base_existe)

@app.route('/processar', methods=['POST'])
def processar():
    if 'file' not in request.files:
        flash("Nenhum arquivo enviado.")
        return redirect(url_for('index'))
        
    file = request.files['file']
    if file.filename == '':
        flash("Selecione um arquivo Excel válido.")
        return redirect(url_for('index'))

    df_rotas = carregar_base_rotas()
    if df_rotas is None:
        flash("Erro: Base de Route Stops não encontrada (verifique a pasta 'bases' ou a raiz).")
        return redirect(url_for('index'))

    try:
        df_bruto = pd.read_excel(file)
        
        col_user_id = encontrar_coluna(df_bruto, ['user_id', 'id_usuario', 'id'])
        col_nome = encontrar_coluna(df_bruto, ['NOME', 'name', 'primeiro_nome', 'passageiro'])
        col_sobrenome = encontrar_coluna(df_bruto, ['SOBRENOME', 'last_name', 'sobrenome_completo'])
        col_endereco = encontrar_coluna(df_bruto, ['ENDEREÇO', 'ENDERECO', 'address', 'logradouro'])
        col_bairro = encontrar_coluna(df_bruto, ['BAIRRO', 'neighborhood'])
        col_cep = encontrar_coluna(df_bruto, ['CEP', 'postal_code', 'zipcode'])
        col_cidade = encontrar_coluna(df_bruto, ['CIDADE', 'city'])
        col_estado = encontrar_coluna(df_bruto, ['ESTADO', 'state', 'uf'])
        col_email = encontrar_coluna(df_bruto, ['email', 'e-mail', 'mail'])
        col_telefone = encontrar_coluna(df_bruto, ['TELEFONE', 'phone_number', 'celular', 'phone', 'tel'])
        col_cpf = encontrar_coluna(df_bruto, ['CPF', 'document', 'documento', 'doc'])
        col_trabalhador = encontrar_coluna(df_bruto, ['ID TRABALHADOR', 'number_worker', 're', 'matricula'])
        col_cargo = encontrar_coluna(df_bruto, ['CARGO', 'position', 'funcao'])
        col_data_inicio = encontrar_coluna(df_bruto, ['DATA INICIO DA RESERVA', 'DATA INÍCIO DA RESERVA', 'DATA INICIO', 'DATA INÍCIO', 'start_date', 'data_inicio'])

        col_rota_entrada = encontrar_coluna(df_bruto, [
            'ROTA ENTRADA', 'ROTA DE ENTRADA', 'NOME ROTA DE ENTRADA', 'rota_entrada', 'route_going',
            'CODIGO ROTA DE ENTRADA', 'CÓDIGO ROTA ENTRADA', 'ROUTE CODE GOING'
        ])
        col_ponto_entrada = encontrar_coluna(df_bruto, [
            'NUMERO PONTO DE ENTRADA', 'NÚMERO PONTO DE ENTRADA', 'PONTO DE ENTRADA', 'ponto_entrada', 'stop_going'
        ])
        col_rota_saida = encontrar_coluna(df_bruto, [
            'ROTA SAIDA', 'ROTA SAÍDA', 'ROTA DE SAIDA', 'ROTA DE SAÍDA', 'NOME ROTA DE SAIDA', 'NOME ROTA DE SAÍDA', 'rota_saida', 'route_return',
            'CODIGO ROTA DE SAIDA', 'CÓDIGO ROTA SAÍDA', 'ROUTE CODE RETURN'
        ])
        col_ponto_saida = encontrar_coluna(df_bruto, [
            'NUMERO PONTO DE SAIDA', 'NÚMERO PONTO DE SAÍDA', 'PONTO DE SAIDA', 'PONTO DE SAÍDA', 'ponto_saida', 'stop_return'
        ])

        rows_final = []

        for _, row in df_bruto.iterrows():
            reg = {}
            
            val_user_id = str(row[col_user_id]).strip() if col_user_id and pd.notna(row[col_user_id]) else ''
            reg['user_id'] = '' if val_user_id.lower() in ['(em branc)', 'nan', 'none'] else val_user_id

            reg['name'] = str(row[col_nome]).strip() if col_nome and pd.notna(row[col_nome]) else ''
            reg['last_name'] = str(row[col_sobrenome]).strip() if col_sobrenome and pd.notna(row[col_sobrenome]) else ''
            reg['address'] = str(row[col_endereco]).strip() if col_endereco and pd.notna(row[col_endereco]) else ''
            reg['address_latitude'] = ''
            reg['address_longitude'] = ''
            reg['neighborhood'] = str(row[col_bairro]).strip() if col_bairro and pd.notna(row[col_bairro]) else ''
            reg['postal_code'] = str(row[col_cep]).strip() if col_cep and pd.notna(row[col_cep]) else ''
            reg['city'] = str(row[col_cidade]).strip() if col_cidade and pd.notna(row[col_cidade]) else ''
            reg['state'] = str(row[col_estado]).strip() if col_estado and pd.notna(row[col_estado]) else ''
            reg['email'] = str(row[col_email]).strip() if col_email and pd.notna(row[col_email]) else ''
            reg['phone_number'] = str(row[col_telefone]).strip() if col_telefone and pd.notna(row[col_telefone]) else ''
            
            doc_val = str(row[col_cpf]).strip() if col_cpf and pd.notna(row[col_cpf]) else ''
            reg['document'] = doc_val

            reg['shifts'] = ''
            reg['cost_center_id'] = ''
            reg['number_worker'] = str(row[col_trabalhador]).strip() if col_trabalhador and pd.notna(row[col_trabalhador]) else ''
            reg['position'] = str(row[col_cargo]).strip() if col_cargo and pd.notna(row[col_cargo]) else ''
            
            val_data = row[col_data_inicio] if col_data_inicio and pd.notna(row[col_data_inicio]) else ''
            if isinstance(val_data, pd.Timestamp):
                reg['start_date'] = val_data.strftime('%d-%m-%Y')
            else:
                reg['start_date'] = str(val_data).strip()

            # Processamento de Entrada (Ida)
            r_going_id = ''
            s_origin_going_id = ''
            company_id_encontrado = ''

            if col_rota_entrada and pd.notna(row[col_rota_entrada]):
                val_r_entrada_norm = normalizar_texto(row[col_rota_entrada])
                df_match_e = pd.DataFrame()

                # 1. Busca por Nome
                if 'Route_name_norm' in df_rotas.columns:
                    df_match_e = df_rotas[df_rotas['Route_name_norm'] == val_r_entrada_norm]

                # 2. Busca por Código (Fallback)
                if df_match_e.empty and 'Route_code_norm' in df_rotas.columns:
                    df_match_e = df_rotas[df_rotas['Route_code_norm'] == val_r_entrada_norm]

                # 3. Busca por correspondência direta bruta
                if df_match_e.empty and 'Route name' in df_rotas.columns:
                    df_match_e = df_rotas[df_rotas['Route name'].astype(str).str.strip() == str(row[col_rota_entrada]).strip()]
                if df_match_e.empty and 'Route code' in df_rotas.columns:
                    df_match_e = df_rotas[df_rotas['Route code'].astype(str).str.strip() == str(row[col_rota_entrada]).strip()]

                if not df_match_e.empty:
                    r_going_id = df_match_e.iloc[0]['Route ID']
                    if 'Site' in df_match_e.columns:
                        company_id_encontrado = obter_company_id(df_match_e.iloc[0]['Site'])
                    
                    if col_ponto_entrada and pd.notna(row[col_ponto_entrada]):
                        try:
                            num_p_e = float(row[col_ponto_entrada])
                            col_p_comp = 'Stop_number_num' if 'Stop_number_num' in df_match_e.columns else 'Stop number'
                            match_p_e = df_match_e[df_match_e[col_p_comp] == num_p_e]
                            if not match_p_e.empty:
                                s_origin_going_id = match_p_e.iloc[0]['Stop ID']
                        except (ValueError, TypeError):
                            pass

            # Processamento de Saída (Volta)
            r_return_id = ''
            s_dest_return_id = ''

            if col_rota_saida and pd.notna(row[col_rota_saida]):
                val_r_saida_norm = normalizar_texto(row[col_rota_saida])
                df_match_s = pd.DataFrame()

                # 1. Busca por Nome
                if 'Route_name_norm' in df_rotas.columns:
                    df_match_s = df_rotas[df_rotas['Route_name_norm'] == val_r_saida_norm]

                # 2. Busca por Código (Fallback)
                if df_match_s.empty and 'Route_code_norm' in df_rotas.columns:
                    df_match_s = df_rotas[df_rotas['Route_code_norm'] == val_r_saida_norm]

                # 3. Busca por correspondência direta bruta
                if df_match_s.empty and 'Route name' in df_rotas.columns:
                    df_match_s = df_rotas[df_rotas['Route name'].astype(str).str.strip() == str(row[col_rota_saida]).strip()]
                if df_match_s.empty and 'Route code' in df_rotas.columns:
                    df_match_s = df_rotas[df_rotas['Route code'].astype(str).str.strip() == str(row[col_rota_saida]).strip()]

                if not df_match_s.empty:
                    r_return_id = df_match_s.iloc[0]['Route ID']
                    if not company_id_encontrado and 'Site' in df_match_s.columns:
                        company_id_encontrado = obter_company_id(df_match_s.iloc[0]['Site'])

                    if col_ponto_saida and pd.notna(row[col_ponto_saida]):
                        try:
                            num_p_s = float(row[col_ponto_saida])
                            col_p_comp = 'Stop_number_num' if 'Stop_number_num' in df_match_s.columns else 'Stop number'
                            match_p_s = df_match_s[df_match_s[col_p_comp] == num_p_s]
                            if not match_p_s.empty:
                                s_dest_return_id = match_p_s.iloc[0]['Stop ID']
                        except (ValueError, TypeError):
                            pass

            # Regra de Senhas
            if not r_going_id and not r_return_id:
                reg['password'] = ''
            elif company_id_encontrado in MELI_COMPANY_IDS:
                reg['password'] = 'Meli@2026'
            else:
                cpf_apenas_numeros = re.sub(r'\D', '', doc_val)
                reg['password'] = f"{cpf_apenas_numeros}@Amz" if cpf_apenas_numeros else ''

            reg['company_id'] = company_id_encontrado
            reg['route_going_id'] = r_going_id
            reg['route_return_id'] = r_return_id
            reg['rate_going_id'] = ''
            reg['rate_return_id'] = ''
            reg['stop_origin_going_id'] = s_origin_going_id
            reg['stop_destination_going_id'] = ''
            reg['stop_origin_return_id'] = ''
            reg['stop_destination_return_id'] = s_dest_return_id
            reg['user_plan_id'] = ''

            rows_final.append(reg)

        df_final = pd.DataFrame(rows_final)[TEMPLATE_COLUMNS]

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Worksheet')
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name='importacao_pronta_busup.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        flash(f"Erro ao processar a planilha: {str(e)}")
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True, port=5000)