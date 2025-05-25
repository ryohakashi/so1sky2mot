# %%
import pandas as pd

listdf_so1sky = pd.read_html('./inbox/1840017534_sales_order-20250516.xls', skiprows=4, header=0)
df_so1sky = listdf_so1sky[0].ffill()
df_so1sky

# %%
df_template = pd.read_table('./template/template_ND6.txt')

df_template

list_template = df_template.to_dict('records')

print(list_template)

# %%
import json

file01 = open("./template/01.json")
json01 = json.loads(file01.read())

print(f"Singkatan: {json01['singkatan']}")

# %%
df_conv = pd.read_json('./template/02.json')
df_conv.set_index("kode", inplace=True)
df_conv

# %%
for index, row in df_so1sky.iterrows():
    # df_data = df_conv[ df_conv['kode'] == row["KD PRODUK"] ]
    
    # df_conv._get_value(row["KD PRODUK"], "kodedist")
    # print(type(row["QTY"]*df_conv._get_value(row["KD PRODUK"], "isi")))
    try:
        if (row["PRODUK ID"] == df_so1sky.iat[index-1, 8]) and (row["AMOUNT"] == 0):
            list_template[len(list_template)-1]["QtyFreeGood"] = row["QTY"] * df_conv.at[row["PRODUK ID"], "isi"]
        else:
            dict_item = {
                    "Nd6tran": "ND6TRAN",
                    "SalesInvoice": "salesinvoice",
                    "Attribute": "value",
                    "SalesmanId": row["SALESMAN ID"],
                    # "SalesOrderNumber": row["ORDER ID"].replace("'", ""),
                    "SalesOrderNumber": row["ORDER ID"],
                    "SalesOrderDate": row["TANGGAL"].replace("-", "/"),
                    # "InvoiceNumber": row["NOMOR"].replace("'", ""),
                    "InvoiceNumber": row["ORDER ID"],
                    "InvoiceDate": row["TANGGAL"].replace("-", "/"),
                    "Term": 0,
                    "SoldToCustomerId": row["CUST ID"].replace(json01['singkatan'], ""),
                    "SentToCustomerId": row["CUST ID"].replace(json01['singkatan'], ""),
                    "InvoicedToCustomerId": row["CUST ID"].replace(json01['singkatan'], ""),
                    "CustomerPo": "",
                    "SellingType": "TO",
                    "DocumentType": "",
                    "CashPayment": 0,
                    "GiroPayment": 0,
                    "GiroNumber": "",
                    "GiroBank": "",
                    "GiroDue": "",
                    "AdjustmentAmount": 0,
                    "Discount1": row["DISKON 1"],
                    "Discount2": row["DISKON 2"],
                    "Discount3": 0,
                    "Tax1": 11,
                    "Tax2": 0,
                    "Tax3": 0,
                    # "ProductCode": row["KD PRODUK"],
                    "ProductCode": df_conv.at[row["PRODUK ID"], "kodedist"],
                    "ProductVarianCode": "-",
                    "QtySold": (row["QTY"] * df_conv.at[row["PRODUK ID"], "isi"]),
                    # "QtySold": (row["QTY"]),
                    "QtyFreeGood": 0,
                    "SellingPrice": row["AMOUNT"],
                    "LineDiscount1": 0,
                    "LineDiscount2": 0,
                    "LineDiscount3": 0,
                    "LineDiscount4": 0,
                    "LineDiscount5": 0,
                    "CompanyId": "NS6022050001143",
                    "BranchId": "1480910000000",
                    "DivisionId": "1481610000000",
                    "WarehouseId": "alokasi",
                    "ManualPonumber": ""
                }
            # print(dict_item)
            list_template.append(dict_item)
            print(row["CUST ID"].replace(json01['singkatan'], ""))
    except KeyError:
        print(f"{row['PRODUK ID']} tidak ditemukan di master harga...!")

# print(list_template)


# %%
df_export = pd.DataFrame.from_records(list_template)

df_export

df_export.to_csv('./outbox/test.txt', sep="|", index=False)


# %%



