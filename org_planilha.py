import pandas as pd
import os

# Correção de dados das estações no INMET (Metricas de ETp)
input_path = '/content/planilhas/estações'
output_path = '/content/planilhas_modificadas'
os.makedirs(output_path, exist_ok=True)

# Loop pelos arquivos
for file_name in os.listdir(input_path):
    if file_name.endswith('.xlsx'):
        file_path = os.path.join(input_path, file_name)
        print(f"\n📄 Processando: {file_name}")

        try:
            excel = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
        except Exception as e:
            print(f"❌ Erro ao abrir {file_name}: {e}")
            continue

        # Procurar aba 'métricas' (case insensitive)
        sheet_name = next((s for s in excel.keys() if s.lower() == 'métricas'), None)

        if not sheet_name:
            print("Aba 'métricas' não encontrada.")
            continue

        df = excel[sheet_name]
        print(f"Aba 'métricas' encontrada: {sheet_name}")
        print(f"Dimensão original: {df.shape}")

        if df.shape[1] < 3:
            print("Menos de 3 colunas. Pulando.")
            continue

        # Coluna C (índice 2)
        col_c = df.iloc[:, 2]

        # Encontrar o primeiro índice onde a coluna C está vazia ou em branco
        cutoff_index = None
        for i, value in enumerate(col_c):
            if pd.isna(value) or str(value).strip() == "":
                cutoff_index = i
                break

        if cutoff_index is None:
            print(" Nenhuma célula vazia encontrada na Coluna C. Nenhuma linha será apagada.")
            df_clean = df.copy()
        else:
            print(f"✂️ Primeira célula vazia na Coluna C encontrada na linha {cutoff_index}")
            df_clean = df.iloc[:cutoff_index].copy()

        # Substituir planilha no dicionário
        excel[sheet_name] = df_clean

        # Salvar arquivo modificado
        output_file_path = os.path.join(output_path, file_name)
        with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
            for sheet, data in excel.items():
                data.to_excel(writer, sheet_name=sheet, index=False)

        print(f"💾 Arquivo salvo: {output_file_path}")

print("\n Todos os arquivos foram processados.")


# Caminho do arquivo zip que será criado
zip_path = '/content/planilhas_modificadas.zip'

# Criar o .zip da pasta de saída
shutil.make_archive(zip_path.replace('.zip', ''), 'zip', output_path)

print(f"\n Arquivo ZIP criado: {zip_path}")