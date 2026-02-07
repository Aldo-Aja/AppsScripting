## Sheets SNBP


### Daya Tampung 2025 (C11)
=IF(Engine!Z2="";""; LET(hasil; IFERROR(XLOOKUP(Engine!Z2;INDIRECT("'"&VLOOKUP(SNBP!C8;Mapping!A:B;2;0)&"'!C8:C");INDIRECT("'"&VLOOKUP(SNBP!C8;Mapping!A:B;2;0)&"'!E8:E"));""); IF(OR(hasil="";hasil=0);"Tidak Ada Data";hasil)))

### Peminat 2024 (C12)
=IF(Engine!Z2="";""; LET(hasil; IFERROR(XLOOKUP(Engine!Z2;INDIRECT("'"&VLOOKUP(SNBP!C8;Mapping!A:B;2;0)&"'!C8:C");INDIRECT("'"&VLOOKUP(SNBP!C8;Mapping!A:B;2;0)&"'!G8:G"));""); IF(OR(hasil="";hasil=0);"Jurusan Baru";hasil)))

### Peminat 2025 (C13)
=IF(Engine!Z2="";""; LET(hasil; IFERROR(XLOOKUP(Engine!Z2;INDIRECT("'"&VLOOKUP(SNBP!C8;Mapping!A:B;2;0)&"'!C8:C");INDIRECT("'"&VLOOKUP(SNBP!C8;Mapping!A:B;2;0)&"'!H8:H"));""); IF(OR(hasil="";hasil=0);"Jurusan Baru";hasil)))

### Rasio keketatan (C14)
=IF(OR(C11="";C13="");""; IF(C13="Jurusan Baru";"Jurusan Baru"; C11/C13))

### Kategori Keketatan (F12)
=IF(Engine!Z2="";""; IF(C14="Jurusan Baru"; "Jurusan Baru"; IFS(C14<0,05;"Sangat Ketat"; C14<0,15;"Ketat"; C14<0,25;"Sedang"; C14<0,4;"Longgar"; TRUE;"Sangat Longgar")))

### Saran penempatan Pilihan (J12)
=IF(
 OR(C9="";F12="");
 "";
 IF(
  OR(
   F12="Sangat Ketat";
   F12="Ketat";
   F12="Sedang"
  );
  "Pilihan 1";
  "Pilihan 1 atau 2 (Fleksibel)"
 )
)

### Rekomendasi Kampus & Jurusan 1 (C17)
=IF(OR(Engine!AF2="";Engine!AF2=0); "❌ Tidak Ditemukan Rekomendasi Jurusan & Kampus"; Engine!AF2)
### Rekomendasi Kampus & Jurusan 2 (C19)
=IF(OR(Engine!AF3="";Engine!AF3=0); "❌ Tidak Ditemukan Rekomendasi Jurusan & Kampus"; Engine!AF3)
### Rekomendasi Kampus & Jurusan 3 (C21)
=IF(OR(Engine!AF4="";Engine!AF4=0); "❌ Tidak Ditemukan Rekomendasi Jurusan & Kampus"; Engine!AF4)
### Rekomendasi Kampus & Jurusan 1 (C22)
=IF(OR(Engine!AF5="";Engine!AF5=0); "❌ Tidak Ditemukan Rekomendasi Jurusan & Kampus"; Engine!AF5)


## Sheets Engine

### Nama Kampus (A1)
=SNBP!C8

### List Jurusan (A2)
=LET(
 ptn; VLOOKUP(SNBP!C8;Mapping!A:B;2;0);
 FILTER(
  INDIRECT("'"&ptn&"'!C8:C");
  INDIRECT("'"&ptn&"'!C8:C")<>"";
  NOT(REGEXMATCH(
    LOWER(INDIRECT("'"&ptn&"'!C8:C"));
    "gelar|daya|minat|portofolio|program studi|snbp"
  ))
 )
)

### Hasil Trim Jurusan (B2)
=TRIM(
 REGEXREPLACE(
  LOWER(
   REGEXREPLACE(
    SNBP!C9;
    "\([^)]*\)";
    ""
   )
  );
  "\b(pendidikan|fakultas|sekolah|program studi|program|studi|guru|kependidikan|ilmu|dan|pengembangan|perencanaan|teknik|sastra|teknologi|kampus|cirebon)\b";
  ""
 )
)

### Lower Nama Jurusan (C2)
=LOWER(
 TRIM(
  REGEXREPLACE(
   SNBP!C9;
   "(^| )(fakultas|sekolah|program studi)( |$)";
   " "
  )
 )
)

### List Nama Kampus yang Menjadi Rekomendasi (D2)
=CHOOSECOLS(FILTER(MASTER_JURUSAN!A2:I; REGEXMATCH(LOWER(MASTER_JURUSAN!C2:C); LOWER(Q2))); 1; 3; 4; 5; 8; 9)

### Nama pendek jurusan dan huruf kecil (P2)
=LOWER(
 TRIM(
  REGEXREPLACE(
   B2;
   "(^| )(pendidikan|sekolah|fakultas|program studi|program|studi|guru|kependidikan|dan|pengembangan|perencanaan|sastra|teknologi|kampus|cirebon)( |$)";
   " "
  )
 )
)

### Hasil Sinonim dari sheets Mapping Sinonim (Q2)
=LET(
 key; Engine!P2;
 keys; SPLIT(LOWER(key);" ");
 data; MappingSinonim!A2:Z;

 cocok;
  FILTER(
   data;
   BYROW(
    data;
    LAMBDA(r;
     SUM(
      BYCOL(
       keys;
       LAMBDA(k;
        COUNTIF(LOWER(r);"*"&k&"*")
       )
      )
     )>0
    )
   )
  );

 IFERROR(
  TEXTJOIN("|"; TRUE; UNIQUE(FLATTEN(cocok)));
  key
 )
)

### Keketatan
=SNBP!C14

### Kode Kampus
=SNBP!M11

### Nama Lengkap Kampus Refrensi cell Z (Z1)
=SNBP!C8

### Nama Jurusan Sebagai Refrensi cell Z (Z2)
=IF(
 SNBP!C8=Engine!Z1;
 SNBP!C9;
 ""
)

### Nama Jurusan Sebagai Refrensi Cell Z dan place Holder Jika Jurusan Kosong (Z3)
=IF(
 Engine!Z2="";
 "Silahkan Pilih Jurusan";
 Engine!Z2
)



