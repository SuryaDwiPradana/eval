from pre_evaluation import *
from post_evaluation import *
from report_evaluation import *

z = 0
while z == 0:
        print("Program Evaluasi Kinerja (BPM): ")
        print("1. Pra - Evaluasi")
        print("2. Post - Evaluasi")
        print("3. Report - Evaluasi")
        print("4. Post & Report - Evaluasi")
        print("5. Bantuan")
        print("6. Keluar")
        x = int(input("Silahkan Pilih: ") or "6")
        if x == 1:
                pre_main()
                z = 1
        elif x == 4:
                post_main()
                report_main()
                z = 1
        elif x == 2:
                post_main()
        elif x == 3:
                report_main()
                z = 1
        elif x == 5:
                print("\n#BANTUAN\n1. Digunakan untuk membuat file excel yang nantinya akan dibagikan ke fungsionaris. (File dibuat berdasarkan isi dari datalist.txt)")
                print("2. Digunakan untuk rekapitulasi hasil dari evaluasi yang diberikan semua fungsionaris.")
                print("3. Digunakan untuk membuat file excel yang nantinya akan diprint. (Nilai akan diambil dari rekapitulasi dari tahap 2)")
                print("4. Proses 3 dan 4 dapat dilakukan secara berkelanjutan")
                print("5. MEnampilkan pesan ini, selamat berevaluasi -RSS-\n")
        elif x == 6:
                z = 1
