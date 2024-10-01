from docx import Document

# membaca file Word
file_path = 'C:/Users/acer/Downloads/Sistem Temu Kembali Informasi/Sentiment Analisis.docx'
doc = Document(file_path)

# mengambil isi document
content = ''
for para in doc.paragraphs:
    content += para.text + '\n'
print()
print("Isi Documentnya :\n",content)
print()


# parsing (Pisahkan isi dokumen jadi beberapa paragraf)
paragraf_list = content.split('\n\n')

# buat list tiap paragraf
paragraf = []  
for i, paragraf_item in enumerate(paragraf_list):
    paragraf.append(paragraf_item)  

print("Paragraf Yang Akan Diolah :\n",paragraf[2])
print()
    
# lexing (menghapus angka)
import re
paragraf[2] = re.sub(r"\d+", "", paragraf[2])
print("Setelah dilakukan lexing (mengapus angka) :\n",paragraf[2])
print()

# lexing (menghilangkan tanda baca)
import string
paragraf[2] = paragraf[2].translate(str.maketrans("","",string.punctuation))
print("Setelah dilakukan lexing (menghilangkan tanda baca) :\n",paragraf[2])
print()

# lexing (menghilangkan karakter selain huruf dan spasi)
paragraf[2] = re.sub(r"[^a-zA-Z\s]", "", paragraf[2])
print("Setelah dilakukan lexing (menghilangkan karakter) :\n",paragraf[2])
print()

# lexing (case folding / mengubh huruf besar menjadi kecil)
paragraf[2] = paragraf[2].lower()
print("Setelah dilakukan lexing (case folding) :\n",paragraf[2])
print()

# lexing (Cleaning (menghapus tautan/url))
paragraf[2] = re.sub(r"http\S+|www\S+|https\S+", "", paragraf[2], flags=re.MULTILINE)
print("Setelah dilakukan lexing (cleaning) :\n",paragraf[2])
print()

# lexing (memisahkan kata per kata)
paragraf[2] = paragraf[2].split()
print("Setelah dilakukan lexing (memisahkan kata per kata) :\n",paragraf[2])
print()

# lexing (Types : menghilangkan kata duplikat)
paragraf[2] = list(set(paragraf[2]))
print("Setelah dilakukan lexing (menghilangkan kata duplikat) :\n",paragraf[2])
print()

# filtering (menghapus kata-kata yang tidak penting)
import nltk
from nltk.corpus import stopwords
stop_words = set(stopwords.words('indonesian'))  # Menggunakan stopwords bahasa Indonesia
paragraf[2] = [word for word in paragraf[2] if word not in stop_words]
print("Setelah dilakukan filtering) :\n",paragraf[2])
print()

# stemming (mengubah bentuk kata menjadi kata dasar)
from Sastrawi.Stemmer.StemmerFactory import StemmerFactory
factory = StemmerFactory()
stemmer = factory.create_stemmer()
paragraf[2] = [stemmer.stem(word) for word in paragraf[2]]
print("Setelah dilakukan stemming :\n",paragraf[2])
print()


'''
saya mempunyai file document yang berisi materi tentang sentiment analysis.
materi tersebut saya ambil dari google untuk dilakukan eksperiment

1) parsing
tujuan dari parsing yaitu memisahkan struktur document dan mengenali komponennya
seperti bagian subject,predikat.
dalam ekperiment saya, saya membagi document menjadi beberapa paragraf
agar mudah diolah.
hanya 1 paragraf yang saya olah ,tujuannya agar saya tau tiap kata akan berubah 
seperti apa tiap langkah

2) lexing 
dalam langkah lexing ini melakukan penghapusan terhadap angka, tanda baca, karakter,
merubah hurus besar menjadi kecil,memisahkan menjadi kata per kata, menghapus elemen
yang tidak dianggap sebagai kata seperti tautan serta menghilangkan kata duplikat

3) filtering
dalam langkah filtering ini melakukan penghapusan terhadap kata yang tidak penting 
dari lexing untuk mewakili document. algoritma yang saya gunakan yaitu algoritma 
stopword

4) stemming
dalam langkah stemming ini mengubah bentuk kata menjadi kata dasar termasuk
menghapus imbuhan dari tiap kata yang akan dijadikan index.
algoritma yang saya gunakan algoritma nazief-adriani yang digunakan khusus
bhs indonesia

*) phrase detection
langkah ini saya tidak dilakukan dalam ekperiment saya karena dari ke 4 langkah 
diatas sudah cukup. selain itu juga butuh komputasinya cukup lama dan kebanyakan
teknik pemrosesan teks tidak menggunakan langkah ini

    Input :
Sentiment analysis merupakan salah satu bidang dari Natural Languange Processing 
(NLP) yang membangun sistem untuk mengenali dan mengekstraksi opini dalam bentuk 
teks. Informasi berbentuk teks saat ini banyak terdapat di internet dalam format 
forum, blog, media sosial, serta situs berisi review. Dengan bantuan sentiment 
analysis, informasi yang tadinya tidak terstruktur dapat diubah menjadi data yang 
lebih terstruktur

    Output :
['nlp', 'natural', 'sosial', 'sistem', 'blog', 'format', 'isi', 'struktur', 'teks'
, 'sentiment', 'ubah', 'languange', 'bentuk', 'media', 'informasi', 'opini', 
'internet', 'bentuk', 'analysis', 'ekstraksi', 'bantu', 'bidang', 'data', 'review'
, 'forum', 'situs', 'bangun', 'nali', 'processing', 'salah']

'''
