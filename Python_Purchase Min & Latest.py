# %%
import imp
import numpy as np
import pandas as pd
import openpyxl
import os 
import xlrd
# Setting Directory
os.chdir('C:/Users/TAM/Documents/Vu Quang Nguyen/VACP/')

# %%
# MUA HANG
URL_mua="../VACP/Export BRAVO/Sổ chi tiết mua hàng (2).xls"
data_muahang=pd.read_excel(URL_mua,usecols="A,C,F,G,I:J",skipfooter=1,header=5).rename(columns={
                                "Unnamed: 0":"Ngày nhập",
                                "Unnamed: 2":"Nhà cung cấp",
                                "Unnamed: 5":"ID",
                                "Unnamed: 6":"Tên hàng",
                                "Số lượng":"Tổng nhập", 
                                "Đơn giá":"Giá nhập"
                                })

#BAN HANG
URL_ban="../VACP/Export BRAVO/Sổ chi tiết bán hàng.xls"  
data_banhang=pd.read_excel(URL_ban,usecols="A,F,K,M",skipfooter=1,header=6).rename(columns={
                                "Unnamed: 5":"ID",
                                "Unnamed: 0":"Ngày bán",
                                })

# %%
data_banhang

# %%
data_muahang

# %%
#Mua hàng giá nhập nhỏ nhất
min_mua_hang=data_muahang.groupby(["ID"])["Giá nhập"].min()
data_muahang_min=data_muahang.merge(min_mua_hang,on="ID",how="left",suffixes=('', '_min'))
data_muahang_min=data_muahang_min[data_muahang_min["Giá nhập"]==data_muahang_min["Giá nhập_min"]
                        ].drop('Giá nhập_min',axis=1).drop_duplicates(subset="ID",keep='first').reset_index(drop=True).rename(
                        columns={
                            "Ngày nhập":"Ngày nhập MIN",
                            "Nhà cung cấp":"Nhà cung cấp MIN",
                            "Tổng nhập":"Tổng nhập MIN",
                            "Giá nhập":"Giá nhập MIN"
                        })



# %%
#Mua hàng gần đây
max_date_muahang=data_muahang.groupby(["ID"])["Ngày nhập"].max()
data_muahang_latest=data_muahang.merge(max_date_muahang,on="ID",how="left",suffixes=('', '_max'))
data_muahang_latest=data_muahang_latest[data_muahang_latest["Ngày nhập"]==data_muahang_latest["Ngày nhập_max"]
                ].drop('Ngày nhập_max',axis=1).drop_duplicates(subset="ID",keep='first').reset_index(drop=True).rename(
                        columns={
                            "Ngày nhập":"Ngày nhập gần đây",
                            "Nhà cung cấp":"Nhà cung cấp gần đây",
                            "Tổng nhập":"Tổng nhập gần đây",
                            "Giá nhập":"Giá nhập gần đây"
                        }
                    )
 

# %%
#Bán hàng gần đây
max_date_banhang=data_banhang.groupby(["ID"])["Ngày bán"].max()
data_banhang_latest=data_banhang.merge(max_date_banhang,on="ID",how="left",suffixes=('', '_max'))
data_banhang_latest=data_banhang_latest[data_banhang_latest["Ngày bán"]==data_banhang_latest["Ngày bán_max"]
                    ].drop('Ngày bán_max',axis=1).drop_duplicates(subset="ID",keep='first').reset_index(drop=True).rename(
                        columns={
                            "Ngày bán":"Ngày bán gần đây",
                            "Giá bán giảm":"Giá bán gần đây"
                        }
                    )


# %%
data_export=data_muahang_latest.merge(data_muahang_min,on='ID',how='left').drop('Tên hàng_y',axis=1).merge(data_banhang_latest,on='ID',how='left').drop_duplicates(subset="ID",keep='first').reset_index(drop=True).rename(
    columns={"Tên hàng_x":"Tên hàng"})
data_export

# %%
# EXPORT FILE DATA
with pd.ExcelWriter('../VACP/Input & Export Campaign/Purchase_Min_Latest.xlsx') as writer:
    data_export.to_excel(writer,sheet_name="Min & Latest")
    # data_report.to_excel(writer,sheet_name="Tổng hợp")
    # data_danhmuchanghoa.to_excel(writer,sheet_name="Danh mục hàng")
    # count_danh_muc.to_excel(writer,sheet_name="đến số lượng")



