import io
from zipfile import ZipFile
import pandas as pd
import glob
import os
import numpy as np
import shutil
import logging
import patoolib as pt

def database():
    try:
        os.makedirs("data")
        tum_ogrenciler = pd.DataFrame([["", "", "", np.NaN, np.NaN, np.NaN, np.NaN, np.NaN, np.NaN, np.NaN, np.NaN, np.NaN, np.NaN, np.NaN, np.NaN, np.NaN, np.NaN, np.NaN, np.NaN, np.NaN, np.NaN, np.NaN, np.NaN, np.NaN, np.NaN]], columns=[
                                      'No', 'Adı', 'Soyadı', 'PC1','PC1a','PC1b', 'PC2', 'PC2a','PC2b','PC3','PC3a','PC3b', 'PC4','PC4a','PC4b', 'PC5','PC5a','PC5b', 'PC6', 'PC7', 'PC8', 'PC9', 'PC10', 'PC11', 'PC1+', 'PC2+', 'PC3+', 'PC4+', 'PC5+', 'PC6+', 'PC7+', 'PC8+', 'PC9+', 'PC10+', 'PC11+'])
        tum_ogrenciler.drop(
            index=tum_ogrenciler.index[0], axis=0, inplace=True)
        completed = pd.Series(dtype="str")
    except:
        tum_ogrenciler = pd.read_pickle(
            os.path.join(os.getcwd()+"/data/students.pkl"))
        completed = pd.read_pickle(os.path.join(
            os.getcwd()+"/data/completed.pkl"))
    return tum_ogrenciler, list(completed)


def store(tum_ogrenciler, completed):
    tum_ogrenciler = pd.DataFrame(tum_ogrenciler)
    tum_ogrenciler.to_pickle(os.path.join(os.getcwd()+"/data/students.pkl"), protocol=4)
    completed = pd.Series(completed)
    completed.to_pickle(os.path.join(os.getcwd()+"/data/completed.pkl"),protocol=4)


def files(folder):
    if folder=="":
        return None
    else:
        filenames = []
        for file in glob.glob(os.path.join(folder+'/**/*.zip'),recursive=True):
            filenames.append(file)
        for file in glob.glob(os.path.join(folder+'/**/*.rar'),recursive=True):
            filenames.append(file)
        for file in glob.glob(os.path.join(folder+'/**/*.xlsx'),recursive=True):
            filenames.append(file)
    return filenames


# def pdffiles(folder):
#     filenames = []
#     for file in glob.glob(os.path.join(folder+"/*.pdf")):
#         filenames.append(file)
#     return filenames


def read_zip(zip_fn):
    zf = ZipFile(zip_fn)
    files = zf.namelist()

    for i in files:
        k=os.path.split(i)[-1]
        if type(k)==list:
            k=k[-1]
        if "PÇ" in k[3:] and k[-4:]=="xlsx":
            target_file = i
            return zf.read(target_file),k

    for i in files:
        k=os.path.split(i)[-1]
        if type(k)==list:
            k=k[-1]
        if "PÇ" in k and k[-4:]=="xlsx":
            target_file = i
            return zf.read(target_file),k
    if not target_file:
        logging.warning(zip_fn,"dosyasında değerlendirme tablosu bulunamadı")
    return None,None

def read_rar(rar_fn):
    temp=os.path.join(os.getcwd()+"/temp")
    try:
        os.mkdir(temp)
    except:
        shutil.rmtree(temp)
        os.mkdir(temp)
    files = []
    pt.extract_archive(rar_fn,outdir=temp)
    for file in glob.glob(temp+'/**/*.xlsx',recursive=True):
        files.append(file)
    for i in files:
        k=os.path.split(i)[-1]
        if type(k)==list:
            k=k[-1]
        if "PÇ" in k[3:] and k[-4:]=="xlsx":
            target_file = i
            return target_file,temp,k

    for i in files:
        k=os.path.split(i)[-1]
        if type(k)==list:
            k=k[-1]
        if "PÇ" in k and k[-4:]=="xlsx":
            target_file = i
            return target_file,temp,k


    if not target_file:
        logging.warning(rar_fn,"dosyasında değerlendirme tablosu bulunamadı")
    return None,None

def get_mezun(filename):
    df = pd.read_excel(filename).iloc[:, 2:5].set_axis(["No", "Soyadı", "Adı"], axis=1)
    return df.dropna()
     

def get_arastirma_proje(filename, completed):
    
    folder_name=os.path.basename(os.path.dirname(filename))+"_"+os.path.basename(filename)
    shortname = "PÇ7, PÇ9, PÇ10 "+ folder_name

    if shortname in completed:
        return "zaten", None

    try:
        df = pd.read_excel(filename).iloc[:, [0,1,13]].set_axis(["No", "AdSoyad", "PC"], axis=1)
        df[['Adı','Soyadı']] = df['AdSoyad'].str.rsplit(n=1,expand=True)
        df.drop("AdSoyad", axis=1, inplace=True)

        df["PC7"] = df.apply(mark_convert, axis=1)
        df["PC9"] = df.apply(mark_convert, axis=1)
        df["PC10"] = df.apply(mark_convert, axis=1)
        df["PC7+"] = 1
        df["PC9+"] = 1
        df["PC10+"] = 1
        completed.append(shortname)
        return df, completed
    
    except:
        return "bilinmeyen", None
    
def get_tasarim_proje(filename, completed):
    folder_name=os.path.basename(os.path.dirname(filename))+"_"+os.path.basename(filename)
    shortname = "PÇ3, PÇ9 "+ folder_name

    if shortname in completed:
        return "zaten", None

    try:
        df = pd.read_excel(filename).iloc[:, [0,1,13]].set_axis(["No", "AdSoyad", "PC"], axis=1)
        df[['Adı','Soyadı']] = df['AdSoyad'].str.rsplit(n=1,expand=True)
        df.drop("AdSoyad", axis=1, inplace=True)

        df["PC3"] = df.apply(mark_convert, axis=1)
        df["PC9"] = df.apply(mark_convert, axis=1)
        df["PC3+"] = 1
        df["PC9+"] = 1
        completed.append(shortname)
        return df, completed
    
    except:
        return "bilinmeyen", None

def mark_convert(row):
    try:
        return (float(row.PC))
    except:
        if row.PC == "AA":
            return 100
        elif row.PC == "BA":
            return 25*3.5
        elif row.PC == "BB":
            return 25*3
        elif row.PC == "CB":
            return 25*2.5
        elif row.PC == "CC":
            return 25*2
        elif row.PC == "DC":
            return 25*1.5
        elif row.PC == "DD":
            return 25*1
        elif row.PC == "FD":
            return 25*0.5
        elif row.PC == "FF":
            return 0
        else:
            return 0

def degerlendirme_notu(file):
    """
    Bazen hocalar değerlendirme skalasını 100 almıyorlar bunun için exceldeki A14
    konumundan skalayı alıp noları çarpmak için hazırlanan bir fonksiyon

    """
    try:
        text=pd.read_excel(file, skiprows=13, usecols="A", nrows=1, header=None, names=["Value"]).iloc[0]["Value"]
        carpan=[int(s) for s in text.split() if s.isdigit()]
        carpan=100/int(carpan[-1])    
        if carpan<2:
            carpan=1        
    except:
        carpan=1
    return carpan


def get_form(filename,tip, completed, sil="Hayır"):
    shortname = os.path.basename(filename)
    completed_name=shortname
    if shortname[-4:]=="xlsx" or shortname[-3:]=="xls" :
        completed_name=os.path.basename(os.path.dirname(filename))+"_"+os.path.basename(filename)

    if completed_name in completed and sil == "Hayır":
        return "zaten", None
    elif completed_name not in completed and sil == "Evet":
        return "Listede zaten yok", "yok"
    try:
        if tip=="zip":
            excel,shortname=read_zip(filename)
            df = pd.read_excel(io.BytesIO(excel), skiprows=20).iloc[:, 1:5].set_axis(
            ["No", "Adı", "Soyadı", "PC"], axis=1)
            carpan=degerlendirme_notu(io.BytesIO(excel))
        elif tip=="rar":
            file,temp,shortname=read_rar(filename)
            df = pd.read_excel(file, skiprows=20).iloc[:, 1:5].set_axis(
            ["No", "Adı", "Soyadı", "PC"], axis=1)
            carpan=degerlendirme_notu(file)
            shutil.rmtree(temp)
        elif tip=="xlsx":
            file=filename
            df = pd.read_excel(file, skiprows=20).iloc[:, 1:5].set_axis(["No", "Adı", "Soyadı", "PC"], axis=1)
            carpan=degerlendirme_notu(file)
            completed_name=os.path.basename(os.path.dirname(filename))+"_"+os.path.basename(filename)
          

        try:
            if type(df["PC"][0])==str:
                df["PC"][0]=0
            df["PC"] = df[["PC"]].apply(pd.to_numeric)
        except:
            df["PC"] = df.apply(mark_convert, axis=1)
            
        df = df[df["No"].str.contains("A")==True]
        PO_of_Lecture = shortname[:4]
        try:
            number = int(PO_of_Lecture[-2:])
        except:
            try:
                number = int(PO_of_Lecture[-2])
            except:
                position=shortname.find("PÇ")
                try:
                    number = int(shortname[position+1:position+3])
                except:
                    number = int(shortname[position+1:position+2])             
        new_name = "PC{}".format(number)
        df = df.rename(columns={"PC": new_name})
        df[new_name]*=carpan
        df[new_name+"+"] = 1
        if sil == "Evet":
            completed.remove(completed_name)
        else:
            completed.append(completed_name)
        return df.dropna(), completed
    except:

        logging.warning("Excel okunamadı: "+str(filename))
        return "bilinmeyen", None


    
    
    
def check_list(filename):
    df = pd.read_excel(filename, skiprows=20).iloc[:, 1:2]
    return df
