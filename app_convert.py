# %%
import json
import pandas as pd

listdf_so1sky = pd.read_html('./inbox/1840017534_sales_order.xls', skiprows=4, header=0)
df_so1sky = listdf_so1sky[0].ffill()
df_so1sky["TANGGAL"] = pd.to_datetime(df_so1sky["TANGGAL"])
print(df_so1sky["TANGGAL"].dtypes)
# df_so1sky

# %%
df_template = pd.read_table('./template/template_ND6.txt')

# df_template

list_template = df_template.to_dict('records')

print(list_template)

# %%


with open("./template/01.json") as file01:
    json01 = json.loads(file01.read())

print(f"Singkatan: {json01['prefix_outlet']}")

# %%
df_conv = pd.read_json('./template/02.json')
df_conv.set_index("kode", inplace=True)
df_conv = df_conv.astype(
    {
        'hirarki': int,
        'uom1': int,
        'uom2': int,
        'uom3': int,
    }
)
print(df_conv.dtypes)

# %%
df_so1sky_merge = pd.merge(df_so1sky, df_conv, how='left', left_on='PRODUK ID', right_on='kode')
print(df_so1sky_merge)

# %%
for index, row in df_so1sky_merge.iterrows():
    # df_data = df_conv[ df_conv['kode'] == row["KD PRODUK"] ]
    
    # df_conv._get_value(row["KD PRODUK"], "kodedist")
    # print(type(row["QTY"]*df_conv._get_value(row["KD PRODUK"], "isi")))
    
    print(f"=>{row['QTY']}")
    print(f"=>{row['HARGA']}")
    match (row['hirarki']):
        case 1 | "1":
            row['QTY'] = row['QTY'] * row['uom1']
            row['HARGA'] = row['HARGA'] / ((100+json01['ppn_order'])/100)
            print("hirarki=>1")
        case 2 | "2":
            row['QTY'] = row['QTY'] * row['uom2']
            row['HARGA'] = row['HARGA'] * (row['uom1'] / row['uom2'])
            row['HARGA'] = row['HARGA'] / ((100+json01['ppn_order'])/100)
        case 3 | "3":
            row['QTY'] = row['QTY'] * row['uom3']
            row['HARGA'] = row['HARGA'] * row['uom1']
            row['HARGA'] = row['HARGA'] / ((100+json01['ppn_order'])/100)
        case _:
            print(row['hirarki'])
    
    print(row)

    if row["AMOUNT"] > 0:
        dict_item = {
            "Nd6tran": "ND6TRAN",
            "SalesInvoice": "salesinvoice",
            "Attribute": "value",
            "SalesmanId": row["SALESMAN ID"],
            # "SalesOrderNumber": row["ORDER ID"].replace("'", ""),
            "SalesOrderNumber": f"{json01['prefix_order']}{row['ORDER ID']}",
            "SalesOrderDate": row["TANGGAL"],
            # "InvoiceNumber": row["NOMOR"].replace("'", ""),
            "InvoiceNumber": f"{json01['prefix_order']}{row['ORDER ID']}",
            "InvoiceDate": row["TANGGAL"],
            "Term": 0,
            "SoldToCustomerId": row["CUST ID"].replace(json01['prefix_outlet'], ""),
            "SentToCustomerId": row["CUST ID"].replace(json01['prefix_outlet'], ""),
            "InvoicedToCustomerId": row["CUST ID"].replace(json01['prefix_outlet'], ""),
            "CustomerPo": "",
            "SellingType": "TO",
            "DocumentType": "",
            "CashPayment": 0,
            "GiroPayment": 0,
            "GiroNumber": "",
            "GiroBank": "",
            "GiroDue": "",
            "AdjustmentAmount": 0,
            "Discount1": 0,
            "Discount2": 0,
            "Discount3": 0,
            "Tax1": 11,
            "Tax2": 0,
            "Tax3": 0,
            # "ProductCode": row["KD PRODUK"],
            "ProductCode": row["kodedist"],
            "ProductVarianCode": "-",
            "QtySold": row["QTY"],
            # "QtySold": (row["QTY"]),
            "QtyFreeGood": 0,
            "SellingPrice": row["HARGA"],
            "LineDiscount1": row["DISKON 1"],
            "LineDiscount2": row["DISKON 2"],
            "LineDiscount3": 0,
            "LineDiscount4": 0,
            "LineDiscount5": 0,
            "CompanyId": "NS6022050001143",
            "BranchId": "1480907626666",
            "DivisionId": "1481611116601",
            "WarehouseId": "alokasi",
            "ManualPonumber": ""
        }
    else:
        # freegood        
        dict_item = {
            "Nd6tran": "ND6TRAN",
            "SalesInvoice": "salesinvoice",
            "Attribute": "value",
            "SalesmanId": row["SALESMAN ID"],
            # "SalesOrderNumber": row["ORDER ID"].replace("'", ""),
            "SalesOrderNumber": f"{json01['prefix_order']}{row['ORDER ID']}",
            "SalesOrderDate": row["TANGGAL"],
            # "InvoiceNumber": row["NOMOR"].replace("'", ""),
            "InvoiceNumber": f"{json01['prefix_order']}{row['ORDER ID']}",
            "InvoiceDate": row["TANGGAL"],
            "Term": 0,
            "SoldToCustomerId": row["CUST ID"].replace(json01['prefix_outlet'], ""),
            "SentToCustomerId": row["CUST ID"].replace(json01['prefix_outlet'], ""),
            "InvoicedToCustomerId": row["CUST ID"].replace(json01['prefix_outlet'], ""),
            "CustomerPo": "",
            "SellingType": "TO",
            "DocumentType": "",
            "CashPayment": 0,
            "GiroPayment": 0,
            "GiroNumber": "",
            "GiroBank": "",
            "GiroDue": "",
            "AdjustmentAmount": 0,
            "Discount1": 0,
            "Discount2": 0,
            "Discount3": 0,
            "Tax1": 11,
            "Tax2": 0,
            "Tax3": 0,
            # "ProductCode": row["KD PRODUK"],
            "ProductCode": row["kodedist"],
            "ProductVarianCode": "-",
            "QtySold": 0,
            # "QtySold": (row["QTY"]),
            "QtyFreeGood": row["QTY"],
            "SellingPrice": row["HARGA"],
            "LineDiscount1": 0,
            "LineDiscount2": 0,
            "LineDiscount3": 0,
            "LineDiscount4": 0,
            "LineDiscount5": 0,
            "CompanyId": "NS6022050001143",
            "BranchId": "1480907626666",
            "DivisionId": "1481611116601",
            "WarehouseId": "alokasi",
            "ManualPonumber": ""
        }

    # print(dict_item)
    list_template.append(dict_item)
    print(row["CUST ID"].replace(json01['prefix_outlet'], ""))

# print(list_template)


# %%
df_export = pd.DataFrame.from_records(list_template)

df_export = df_export.groupby(
    [
        "Nd6tran",
        "SalesInvoice",
        "Attribute",
        "SalesmanId",
        "SalesOrderNumber",
        "SalesOrderDate",
        "InvoiceNumber",
        "InvoiceDate",
        "Term",
        "SoldToCustomerId",
        "SentToCustomerId",
        "InvoicedToCustomerId",
        "CustomerPo",
        "SellingType",
        "DocumentType",
        "CashPayment",
        "GiroPayment",
        "GiroNumber",
        "GiroBank",
        "GiroDue",
        "AdjustmentAmount",        
        "Tax1",
        "Tax2",
        "Tax3",
        "ProductCode",
        "ProductVarianCode",
        "SellingPrice",
        "LineDiscount1",
        "LineDiscount2",
        "LineDiscount3",
        "LineDiscount4",
        "LineDiscount5",
        "CompanyId",
        "BranchId",
        "DivisionId",
        "WarehouseId",
        "ManualPonumber"
    ]
).agg(
    {
        "Discount1": "sum",
        "Discount2": "sum",
        "Discount3": "sum",
        "QtySold": "sum", 
        "QtyFreeGood": "sum"
    }
).reset_index()

df_export = df_export.reindex(columns=
    [
        "Nd6tran",
        "SalesInvoice",
        "Attribute",
        "SalesmanId",
        "SalesOrderNumber",
        "SalesOrderDate",
        "InvoiceNumber",
        "InvoiceDate",
        "Term",
        "SoldToCustomerId",
        "SentToCustomerId",
        "InvoicedToCustomerId",
        "CustomerPo",
        "SellingType",
        "DocumentType",
        "CashPayment",
        "GiroPayment",
        "GiroNumber",
        "GiroBank",
        "GiroDue",
        "AdjustmentAmount",
        "Discount1",
        "Discount2",
        "Discount3",
        "Tax1",
        "Tax2",
        "Tax3",
        "ProductCode",
        "ProductVarianCode",
        "QtySold", 
        "QtyFreeGood",
        "SellingPrice",
        "LineDiscount1",
        "LineDiscount2",
        "LineDiscount3",
        "LineDiscount4",
        "LineDiscount5",
        "CompanyId",
        "BranchId",
        "DivisionId",
        "WarehouseId",
        "ManualPonumber"
    ]
)

df_export.to_csv('./outbox/test.txt', sep="|", index=False)


# %%



