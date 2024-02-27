# import os
# def install_and_import():
#     try:
#         import numpy as np
#         import pandas as pd
#         import matplotlib.pyplot as plt
#         import openpyxl
#     except ImportError:
#         import pip
#         pip.main(['install', 'pandas'])
#         pip.main(['install', 'matplotlib'])
#         pip.main(['install', 'openpyxl'])
# install_and_import()
# os.system("cls")

import pandas as pd
pd.set_option('display.max_colwidth', None)
import matplotlib.pyplot as plt
plt.figure(figsize=(14, 8))

###########   CÂU B
print("\n                                  DANH SÁCH SINH VIÊN")
df_sinh_vien = pd.read_excel('DuLieuThucHanh2_K66_GuiSV_V1.xlsx', sheet_name="Sinh Vien")
df_sinh_vien['STT'] = [i for i in range(1, len(df_sinh_vien) + 1)]
print(df_sinh_vien.to_string(index=False))

print("\n                           DANH SÁCH ĐIỂM MÔN 1")
df_diem_mon_1 = pd.read_excel('DuLieuThucHanh2_K66_GuiSV_V1.xlsx', sheet_name="Diem Mon 1")
df_diem_mon_1['STT'] = [i for i in range(1, len(df_diem_mon_1) + 1)]
print(df_diem_mon_1.to_string(index=False))

print("\n                           DANH SÁCH ĐIỂM MÔN 2")
df_diem_mon_2 = pd.read_excel('DuLieuThucHanh2_K66_GuiSV_V1.xlsx', sheet_name="Diem Mon 2")
df_diem_mon_2['STT'] = [i for i in range(1, len(df_diem_mon_2) + 1)]
print(df_diem_mon_2.to_string(index=False))

print("\n                           DANH SÁCH ĐIỂM MÔN 3")
df_diem_mon_3 = pd.read_excel('DuLieuThucHanh2_K66_GuiSV_V1.xlsx', sheet_name="Diem Mon 3")
df_diem_mon_3['STT'] = [i for i in range(1, len(df_diem_mon_3) + 1)]
print(df_diem_mon_3.to_string(index=False))

########### CÂU C
print("\n\n                                            DANH SÁCH ĐIỂM MÔN TỔNG HỢP")

#Merge bảng 1
df_tong_hop = pd.merge(df_sinh_vien, df_diem_mon_1, on="Ma Sinh Vien", how="inner").T.drop_duplicates().T
df_tong_hop.rename(columns={'Diem Qua Trinh': 'Diem Qua Trinh Mon 1'}, inplace=True)
df_tong_hop.rename(columns={'Diem Cuoi Ky': 'Diem Cuoi Ky Mon 1'}, inplace=True)
df_tong_hop['Diem Tong Ket Mon 1'] = df_tong_hop['Diem Qua Trinh Mon 1'] * 0.3 + df_tong_hop['Diem Cuoi Ky Mon 1'] * 0.7

#Merge bảng 2
df_tong_hop = pd.merge(df_tong_hop, df_diem_mon_2, on="Ma Sinh Vien", how="inner").drop(columns=['STT', 'Ho Ten'])
df_tong_hop.rename(columns={'Diem Qua Trinh': 'Diem Qua Trinh Mon 2'}, inplace=True)
df_tong_hop.rename(columns={'Diem Cuoi Ky': 'Diem Cuoi Ky Mon 2'}, inplace=True)
df_tong_hop['Diem Tong Ket Mon 2'] = df_tong_hop['Diem Qua Trinh Mon 2'] * 0.3 + df_tong_hop['Diem Cuoi Ky Mon 2'] * 0.7

#Merge bảng 3
df_tong_hop = pd.merge(df_tong_hop, df_diem_mon_3, on="Ma Sinh Vien", how="inner").drop(columns=['STT', 'Ho Ten'])
df_tong_hop.rename(columns={'Diem Qua Trinh': 'Diem Qua Trinh Mon 3'}, inplace=True)
df_tong_hop.rename(columns={'Diem Cuoi Ky': 'Diem Cuoi Ky Mon 3'}, inplace=True)
df_tong_hop.rename(columns={'STT_x': 'STT'}, inplace=True)
df_tong_hop.rename(columns={'Ho Ten_x': 'Ho Va Ten'}, inplace=True)
df_tong_hop['Diem Tong Ket Mon 3'] = df_tong_hop['Diem Qua Trinh Mon 3'] * 0.3 + df_tong_hop['Diem Cuoi Ky Mon 3'] * 0.7

df_tong_hop['Diem Trung Binh'] = (df_tong_hop['Diem Tong Ket Mon 1'] + df_tong_hop['Diem Tong Ket Mon 2'] + df_tong_hop['Diem Tong Ket Mon 3'])/3
#df_tong_hop['Diem Trung Binh'] = df_tong_hop['Diem Trung Binh'].round(2)
df_tong_hop['STT'] = [i for i in range(1, len(df_tong_hop) + 1)]

print(df_tong_hop.to_string(index=False))

#########   CÂU D
print("\n\n                                            DANH SÁCH SINH VIÊN TRƯỢT MÔN THỨ NHẤT")
df_sinh_vien_khong_qua_mon_1 = df_tong_hop[df_tong_hop['Diem Tong Ket Mon 1'] < 4.0][['STT',
                                                                                      'Ma Sinh Vien',
                                                                                      'Ho Va Ten',
                                                                                      'Ngay Sinh',
                                                                                      'Gioi Tinh',
                                                                                      'Lop Quan Ly',
                                                                                      'Diem Qua Trinh Mon 1',
                                                                                      'Diem Cuoi Ky Mon 1',
                                                                                      'Diem Tong Ket Mon 1']]
df_sinh_vien_khong_qua_mon_1['STT'] = [i for i in range(1, len(df_sinh_vien_khong_qua_mon_1) + 1)]
print(df_sinh_vien_khong_qua_mon_1.to_string(index=False))


print("\n\n                                            DANH SÁCH SINH VIÊN TRƯỢT MÔN THỨ HAI")
df_sinh_vien_khong_qua_mon_2 = df_tong_hop[df_tong_hop['Diem Tong Ket Mon 2'] < 4.0][['STT',
                                                                                      'Ma Sinh Vien',
                                                                                      'Ho Va Ten',
                                                                                      'Ngay Sinh',
                                                                                      'Gioi Tinh',
                                                                                      'Lop Quan Ly',
                                                                                      'Diem Qua Trinh Mon 2',
                                                                                      'Diem Cuoi Ky Mon 2',
                                                                                      'Diem Tong Ket Mon 2']]
df_sinh_vien_khong_qua_mon_2['STT'] = [i for i in range(1, len(df_sinh_vien_khong_qua_mon_2) + 1)]
print(df_sinh_vien_khong_qua_mon_2.to_string(index=False))

print("\n\n                                            DANH SÁCH SINH VIÊN TRƯỢT MÔN THỨ BA")
df_sinh_vien_khong_qua_mon_3 = df_tong_hop[df_tong_hop['Diem Tong Ket Mon 3'] < 4.0][['STT',
                                                                                      'Ma Sinh Vien',
                                                                                      'Ho Va Ten',
                                                                                      'Ngay Sinh',
                                                                                      'Gioi Tinh',
                                                                                      'Lop Quan Ly',
                                                                                      'Diem Qua Trinh Mon 3',
                                                                                      'Diem Cuoi Ky Mon 3',
                                                                                      'Diem Tong Ket Mon 3']]
df_sinh_vien_khong_qua_mon_3['STT'] = [i for i in range(1, len(df_sinh_vien_khong_qua_mon_3) + 1)]
print(df_sinh_vien_khong_qua_mon_3.to_string(index=False))


#########   CÂU E

print("\n\n                                            DANH SÁCH SINH VIÊN KHÔNG TRƯỢT NÀO")

df_sinh_vien_khong_truot_mon_nao = df_tong_hop[(df_tong_hop['Diem Tong Ket Mon 1'] >= 4.0) &
                                               (df_tong_hop['Diem Tong Ket Mon 2'] >= 4.0) &
                                               (df_tong_hop['Diem Tong Ket Mon 3'] >= 4.0)][['STT',
                                                                                      'Ma Sinh Vien',
                                                                                      'Ho Va Ten',
                                                                                      'Ngay Sinh',
                                                                                      'Gioi Tinh',
                                                                                      'Lop Quan Ly',
                                                                                      'Diem Tong Ket Mon 1',
                                                                                      'Diem Tong Ket Mon 2',
                                                                                      'Diem Tong Ket Mon 3',
                                                                                      'Diem Trung Binh']]
df_sinh_vien_khong_truot_mon_nao['STT'] = [i for i in range(1, len(df_sinh_vien_khong_truot_mon_nao) + 1)]
print(df_sinh_vien_khong_truot_mon_nao.to_string(index=False))

#  CÂU F

figure, axis = plt.subplots(2, 2)
figure.set_size_inches(20, 11.25)
axis[1, 1].barh(df_tong_hop['Ho Va Ten'], df_tong_hop['Diem Trung Binh'], height=0.5, color='g')
axis[1, 1].set_title('Điểm Trung Bình')
axis[0, 0].barh(df_tong_hop['Ho Va Ten'], df_tong_hop['Diem Tong Ket Mon 1'], height=0.5, color='m')
axis[0, 0].set_title('Điểm Môn 1')
axis[0, 1].barh(df_tong_hop['Ho Va Ten'], df_tong_hop['Diem Tong Ket Mon 2'], height=0.5, color='c')
axis[0, 1].set_title('Điểm Môn 2')
axis[1, 0].barh(df_tong_hop['Ho Va Ten'], df_tong_hop['Diem Tong Ket Mon 3'], height=0.5)
axis[1, 0].set_title('Điểm Môn 3')
plt.tight_layout(pad=0.4, w_pad=0.5, h_pad=1.0)
plt.savefig('Thống kê điểm.png')
plt.show()


#   CÂU G
with pd.ExcelWriter('[Nhóm 12]_Danh-Sách-Sinh-Viên-Final.xlsx') as writer:
    df_sinh_vien.to_excel(writer, sheet_name='DS Sinh Viên')
    df_diem_mon_1.to_excel(writer, sheet_name='DS Điểm Môn Thứ 1')
    df_sinh_vien_khong_qua_mon_1.to_excel(writer, sheet_name='DS Trượt Môn Thứ 1')
    df_diem_mon_2.to_excel(writer, sheet_name='DS Điểm Môn Thứ 2')
    df_sinh_vien_khong_qua_mon_2.to_excel(writer, sheet_name='DS Trượt Môn Thứ 2')
    df_diem_mon_3.to_excel(writer, sheet_name='DS Điểm Môn Thứ 3')
    df_sinh_vien_khong_qua_mon_3.to_excel(writer, sheet_name='DS Trượt Môn Thứ 3')
    df_sinh_vien_khong_truot_mon_nao.to_excel(writer, sheet_name='DS Sinh Viên Không Trượt Môn')
    df_tong_hop.to_excel(writer, sheet_name='DS Tổng Hợp')