import io
import pandas as pd

# pip install xlsxwriter

dict1 = [{
    "name": "Jane Smith",
    "address": "123 Maple Street",
    "city": "Pretendville",
    "state": "NY",
    "zip": "12345"}]

dict2 = [{
    "name": "Lebrom James",
    "address": "123 Maple Street",
    "city": "Pretendville",
    "state": "NY",
    "zip": "12345"}]

dict3 = [{
    "name": "Anthony Davis",
    "address": "123 Maple Street",
    "city": "Pretendville",
    "state": "NY",
    "zip": "12345"}]

try:
    dict1_df = pd.DataFrame(dict1)
    dict2_df = pd.DataFrame(dict2)
    dict3_df = pd.DataFrame(dict3)

    with io.BytesIO() as output:
        with pd.ExcelWriter(output, engine='xlsxwriter') as ew:
            dict1_df.to_excel(ew, sheet_name='Perguntas', index=False)
            dict2_df.to_excel(ew, sheet_name='Politicas', index=False)
            dict3_df.to_excel(ew, sheet_name='Recurso', index=False)

        data = output.getvalue()

except Exception as e:
    print("Erro na montagem dos bytes")
