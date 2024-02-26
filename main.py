import os
def install_and_import():
    try:
        import numpy as np
        import pandas as pd
        import matplotlib.pyplot as plt
        import openpyxl
    except ImportError:
        import pip
        pip.main(['install', 'pandas'])
        pip.main(['install', 'numpy'])
        pip.main(['install', 'matplotlib'])
        pip.main(['install', 'openpyxl'])
install_and_import()
os.system("cls")
import numpy as np
import pandas as pd
pd.set_option('display.max_colwidth', None)
import matplotlib.pyplot as plt
plt.figure(figsize=(14, 8))

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

#BIỂU ĐỒ CỘT DỌC
#colors = ['red', 'black', 'blue', 'green', 'blue', 'green', 'blue', 'green', 'blue', 'green', 'blue', 'green', 'blue', 'green', 'blue', 'green', 'blue', 'green', 'blue', 'green', 'blue', 'green', 'blue', 'green', 'blue', 'green', 'blue']
# plt.bar(df_tong_hop['Ho Va Ten'], df_tong_hop['Diem Trung Binh'], color='orange')
# #plt.xlabel('Ho Va Ten')
# plt.ylabel('Điểm Trung Bình')
# plt.title('Điểm Trung Bình của các sinh viên')
# plt.xticks(rotation=90)
# plt.tight_layout()
# plt.show()

#########BIỂU ĐỒ CỘT KẾT HỢP
# x = np.arange(0, len(df_tong_hop)*3, 3)
# plt.barh(x - 0.5, df_tong_hop['Diem Trung Binh'], height=0.5)
# plt.barh(x, df_tong_hop['Diem Tong Ket Mon 1'], height=0.5)
# plt.barh(x + 0.5, df_tong_hop['Diem Tong Ket Mon 2'], height=0.5)
# plt.barh(x + 1, df_tong_hop['Diem Tong Ket Mon 3'], height=0.5)
#
# plt.xlabel('Điểm Trung Bình')
# plt.title('Điểm Trung Bình của các sinh viên')
# plt.yticks(x, df_tong_hop['Ho Va Ten'])
# plt.tight_layout()
# plt.legend(['Điểm Môn 1', 'Điểm Môn 2', 'Điểm Môn 3', 'Điểm Trung Bình'][::-1])
# plt.show()


#BIỂU ĐỒ CỘT NGANG
# plt.barh(df_tong_hop['Ho Va Ten'], df_tong_hop['Diem Trung Binh'], height=0.5)
# plt.xlabel('Điểm Trung Bình')
# plt.title('Điểm Trung Bình của các sinh viên')
# plt.tight_layout()
# plt.show()


#BIỂU ĐỒ ĐƯỜNG
# for i in range(len(df_tong_hop)):
#     plt.plot(['Môn Thứ Nhất', 'Môn Thứ Hai', 'Môn Thứ Ba', 'Trung Bình'], [df_tong_hop['Diem Tong Ket Mon 1'][i], df_tong_hop['Diem Tong Ket Mon 2'][i], df_tong_hop['Diem Tong Ket Mon 3'][i], df_tong_hop['Diem Trung Binh'][i]], label=df_tong_hop['Ho Va Ten'][i])
# # plt.plot([df_tong_hop['Diem Tong Ket Mon 1'], df_tong_hop['Diem Tong Ket Mon 2'],
# #           df_tong_hop['Diem Tong Ket Mon 3'], df_tong_hop['Diem Trung Binh']], [df_tong_hop['Ho Va Ten']])
# plt.legend(loc='lower right', prop={'size': 6}, bbox_to_anchor=(1.13, 0))
#
# plt.show()

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