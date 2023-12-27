# Szakdolgozat 2023 Farkas Krisztina

import os, shutil
import sys
import exiftool
from zipfile import ZipFile
from exiftool import ExifToolHelper


# Ellenőrizni a három paraméter és a két hivatkozott folder létezését

#i = 0
#while i < n:
#    print(sys.argv[i])
#    i += 1


def write_to_csv(xpath,filename,data):
    # data beírása csv-be, file megnyitás: append, egy sor hozzáírása
    # az input formátuma miatt : cseréje vesszőre!!!!!
    #data2 = data.replace(":",";")
    #print("CSV: ",data);
    csvpath = os.path.join(sys.argv[2], sys.argv[3])
    file_csv = open(csvpath,"a")
    file_csv.writelines(data+"\n")
    file_csv.close()
    return{}

# speciális kereső, a találat megy a CSV-be
def search_to_csv(fn,xfrom, xtext,xto, csvpath):
    with open(fn, 'r', encoding="utf-8") as file:
        content: str = file.read()
        xi = content.find(xfrom)
        if xi != -1:
            xfrom_len = len(xfrom)
            xj = content.find(xto)
            if xj != -1:
                xto_len = len(xto)
                yfrom = xi+xfrom_len+1
                xcont = content[xi+xfrom_len:xj]
                write_to_csv(sys.argv[2],sys.argv[3],f" {xtext};{xcont}")

    return {}

# speciális kereső rsidRoot-ra, a találat megy a CSV-be
def search_to_csv2(fn,xfrom, xtext, csvpath):
    with open(fn, 'r', encoding="utf-8") as file:
        content: str = file.read()
        xi = content.find(xfrom)
        if xi != -1:
            xfrom_len = len(xfrom)
            yfrom = xi+xfrom_len+1
            xcont = content[xi+xfrom_len:xi+xfrom_len+8]
            write_to_csv(sys.argv[2],sys.argv[3],f" {xtext};{xcont}")

    return {}

# speciális kereső rsidRoot-ra, a találat megy a CSV-be
def search_to_csv3(fn,xfrom, xtext, csvpath):
    with open(fn, 'r', encoding="utf-8") as file:
        content: str = file.read()
        xcount = content.count(xfrom)
        write_to_csv(sys.argv[2],sys.argv[3],f" {xtext};{xcount}")

    return {}

def read_meta_jpg(fn):
    data = ""
    # filenév elé path
    file = sys.argv[1]+"\\"+fn
    # metaadatok kiszedése
    with ExifToolHelper() as et:
        # Összes metaadat, megnevezéssel
        for d in et.get_metadata(file):
            # egyenkénti listázás
            for k, v in d.items():
                # Kettőspont cseréje ;-re, CSV miatt
                k2 = k.replace(":",";")
                write_to_csv(sys.argv[2],sys.argv[3],f" {k2};{v}")

    return{}

def read_meta_excel(fn):
    # lista [file],[megnevezés],[kezdete],[vége] #### helyette spéci kereső
    lista_app_path = 'docProps\\app.xml'
    lista_core_path = 'docProps\\core.xml'

    # tmp folder létrehozása
    path = os.path.join(sys.argv[2],"temp")
    isdir = os.path.exists(path)
    if (not isdir):
        os.mkdir(path)

    # UNZIP
    pathzip = os.path.join(sys.argv[1],fn)
    with ZipFile(pathzip) as zObject:
        zObject.extractall(path)

    # SourceFile
    write_to_csv(sys.argv[2], sys.argv[3], f" SourceFile;{fn}")

    # Keresések
    path = os.path.join(sys.argv[2], "temp\\docProps\\app.xml")
    csvpath = os.path.join(sys.argv[2], sys.argv[3])
    search_to_csv(path,'<Application>','Application','</Application>',csvpath)
    search_to_csv(path,'<AppVersion>','AppVersion','</AppVersion>',csvpath)

    path = os.path.join(sys.argv[2], "temp\\docProps\\core.xml")
    csvpath = os.path.join(sys.argv[2], sys.argv[3])
    search_to_csv(path,'<dc:creator>','Creator','</dc:creator>',csvpath)
    search_to_csv(path,'<cp:lastModifiedBy>','LastModifiedBy','</cp:lastModifiedBy>',csvpath)
    search_to_csv(path,'<dcterms:created xsi:type="dcterms:W3CDTF">','Created','</dcterms:created>',csvpath)
    search_to_csv(path,'<dcterms:modified xsi:type="dcterms:W3CDTF">','Modified','</dcterms:modified>',csvpath)

    # tmp folder törlése
    path = os.path.join(sys.argv[2], "temp")
    if os.path.isdir(path):
        shutil.rmtree(path)

    return {}

def read_meta_word(fn):
    # lista [file],[megnevezés],[kezdete],[vége] #### helyette spéci kereső
    lista_app_path = 'docProps\\app.xml'
    lista_core_path = 'docProps\\core.xml'
    lista_settings_path = "word\\"+'settings.xml'
    lista_settings =[['rsidRoot','<w:rsidRoot w:val="']]
    # Az rsid számlálást külön kell megoldani

    # tmp folder létrehozása
    path = os.path.join(sys.argv[2],"temp")
    isdir = os.path.exists(path)
    if (not isdir):
        os.mkdir(path)

    # UNZIP
    pathzip = os.path.join(sys.argv[1],fn)
    with ZipFile(pathzip) as zObject:
        zObject.extractall(path)

    # SourceFile
    write_to_csv(sys.argv[2], sys.argv[3], f" SourceFile;{fn}")

    # Keresések
    path = os.path.join(sys.argv[2], "temp\\docProps\\app.xml")
    csvpath = os.path.join(sys.argv[2], sys.argv[3])
    search_to_csv(path,'<TotalTime>','TotalTime','</TotalTime>',csvpath)
    search_to_csv(path,'<Pages>','Pages','</Pages>',csvpath)
    search_to_csv(path,'<Words>','Words','</Words>',csvpath)
    search_to_csv(path,'<Characters>','Characters','</Characters>',csvpath)
    search_to_csv(path,'<CharactersWithSpaces>','CharactersWithSpaces','</CharactersWithSpaces>',csvpath)
    search_to_csv(path,'<Lines>','Lines','</Lines>',csvpath)
    search_to_csv(path,'<Paragraphs>','Paragraphs','</Paragraphs>',csvpath)
    search_to_csv(path,'<Application>','Application','</Application>',csvpath)
    search_to_csv(path,'<AppVersion>','AppVersion','</AppVersion>',csvpath)

    path = os.path.join(sys.argv[2], "temp\\docProps\\core.xml")
    csvpath = os.path.join(sys.argv[2], sys.argv[3])
    search_to_csv(path,'<dc:title>','Title','</dc:title>',csvpath)
    search_to_csv(path,'<dc:subject>','Subject','</dc:subject>',csvpath)
    search_to_csv(path,'<dc:creator>','Creator','</dc:creator>',csvpath)
    search_to_csv(path,'<cp:lastModifiedBy>','LastModifiedBy','</cp:lastModifiedBy>',csvpath)
    search_to_csv(path,'<dcterms:created xsi:type="dcterms:W3CDTF">','Created','</dcterms:created>',csvpath)
    search_to_csv(path,'<dcterms:modified xsi:type="dcterms:W3CDTF">','Modified','</dcterms:modified>',csvpath)

    path = os.path.join(sys.argv[2], "temp\\word\\settings.xml")
    csvpath = os.path.join(sys.argv[2], sys.argv[3])
    search_to_csv2(path,'<w:rsidRoot w:val="','rsidRoot',csvpath) # speciális rsidRoot kereső, mert a záró tag máshol is előfordul, viszont fix hosszú a tartalom
    search_to_csv3(path,'<w:rsid w:val="','rsid count',csvpath) # speciális rsid számláló

    # tmp folder törlése
    path = os.path.join(sys.argv[2], "temp")
    if os.path.isdir(path):
        shutil.rmtree(path)
    return{}


def read_files(xpath):
    # könyvtárak ellenőrzése
    err = 0
    path = sys.argv[1]
    isdir = os.path.isdir(path)
    if (not isdir):
        print(f"{path} könyvtár nem létezik")
        err = 1
    path = sys.argv[2]
    isdir = os.path.isdir(path)
    if (not isdir):
        print(f"{path} könyvtár nem létezik")
        err = 1
    if (err == 0):
        # itt a függvény
        dir_list = os.listdir(xpath) # dir_list tömb lesz, tartalma a file lista
        i = 0
        n = len(dir_list)
        while i < n:
            # típus megállapítása és a megfelelő function hívása
            st = dir_list[i] # beteszem a köv. file nevét
            ext_start = st.rindex(".") # pont utolsó előfordulása
            ext = st[ext_start+1:]
            #print(st)
            #print(ext)
            #print("")
            # CASE struktúra
            match ext:
                case "jpg":
                    read_meta_jpg(st)
                case "docx":
                    read_meta_word(st)
                case "xlsx":
                    read_meta_excel(st)
                case _:
                    print("Nem megfelelő kiterjesztés ",st) # nem feltétlenül szükséges kiírni, csak infó

            i += 1

# demó függvény használata
if __name__ == '__main__':
    # Input paraméterek: input path; output path; output file neve
    n = len(sys.argv)
    if n != 4:
        print ('A paraméterek száma nem megfelelő')
    else:
        read_files(sys.argv[1])
