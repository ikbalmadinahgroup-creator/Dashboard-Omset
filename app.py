"""Dashboard Omset MFlash

Dashboard Streamlit untuk memantau Omset, Iklan (Meta Ads), Walk-in,
6 Pilar MFlash, Kontribusi Marketing Corporate vs Sales Retail, dan
Project Tracker Sales & Marketing di 18 cabang MFlash. Termasuk styling
tabel Walk-in (kotak + warna), export tabel Walk-in & Scoreboard ke
JPG/PDF, insight otomatis, dan export laporan lengkap ke PPTX/PDF.
Scoreboard mengikuti persis format & rumus pada sheet "Scoreboard" di
file Excel master (target kuartalan, EXPECTED VALUE berdasar hari
berjalan dalam kuartal, % PENCAPAIAN = S/D HARI INI dibagi EXPECTED
VALUE). Target Omset bersifat KUARTALAN (berlaku 1 kuartal penuh - mis.
Jul-Sep - dan TIDAK perlu diupload ulang tiap hari; hanya perlu upload
ulang saat masuk kuartal berikutnya dengan angka target yang berbeda).
Loader Target mendukung DUA format file: format panjang (kolom Cabang,
Kategori, Target - satu baris per kombinasi cabang x kategori, sesuai
template bawaan) DAN format lebar (kolom Cabang, Target Service, Target
Gadget, Target All, dst - satu baris per cabang dengan semua kategori
sekaligus, format yang lebih umum dipakai user secara natural).
Tabel Walk-in per cabang bersifat KUMULATIF dari awal kuartal (1 Juli/
Okt/Jan/Apr) sampai Tanggal Acuan yang dipilih (konsisten dengan S/D
HARI INI di Scoreboard). Loader data Omset & Walk-in memakai
pandas.read_excel (bukan openpyxl read_only) supaya tahan terhadap
file export MFlash dengan metadata dimensi sheet yang tidak akurat.
Loader Iklan mem-buang kolom duplikat sebelum digabung (pd.concat)
untuk mencegah pandas.errors.InvalidIndexError. Tab Sales & Marketing
berisi project tracker interaktif (tambah/edit/hapus baris langsung di
dashboard, dibungkus st.form supaya mengetik di tabel tidak memicu
rerun seluruh dashboard) dengan status, due date, PIC, progress (+
kendala, catatan, action plan per project), dan ringkasan progress
berwarna per status di bawah tabel (karena ProgressColumn bawaan
Streamlit tidak mendukung warna kustom).
Loader per-file di-cache (st.cache_data, key = path+mtime+size) supaya
file Excel yang belum berubah tidak dibaca ulang setiap kali ada
interaksi di dashboard (setiap klik/filter membuat Streamlit menjalankan
ulang seluruh script). Selain itu, hasil gabungan semua file Omset/Iklan/
Walk-in juga disimpan sebagai cache parquet di disk (+ backup ke GitHub
kalau aktif) yang tetap ada walau aplikasi baru saja restart/"bangun
tidur" di Streamlit Cloud — ini fix utama untuk keluhan loading lambat
saat cold-start, karena tanpa cache ini SEMUA file Excel per cabang harus
dibaca ulang dari nol tiap kali proses Streamlit baru dimulai (bisa
puluhan detik untuk 18 cabang), padahal baca cache parquet hanya makan
waktu di bawah 0,1 detik selama file Excel sumber belum berubah.
Log riwayat & backup GitHub juga hanya jalan saat data benar-benar
berubah, bukan di setiap rerun.
"""

import base64
import calendar
import io
import os
import re
from datetime import datetime, date, timedelta

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

import openpyxl
import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st
from PIL import Image

st.set_page_config(page_title="Dashboard Omset MFlash", page_icon="📊", layout="wide")

LOGO_BASE64 = "iVBORw0KGgoAAAANSUhEUgAAASwAAADUCAYAAAAmyx61AAAuOElEQVR4nO3de3xV1Zk38N/zrL3PyUlCIDcuigIB0SLiJQlQGYvW1mI7fWvbwUoSsLaOTm2tCt6mtkOZttrqCFqr0zq1rUJAzbS+tdV2pq2XvtYCId6LlqsoipAb5HZyztlrPe8fJ8EASUhCknOQ5/v5xA+es7PXs5N9nqzbXgtQSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSimllFJKKaWUOtpRqgNQKTR3rjd2X1GuNW05jjzPM16s3U/s2/fc6sZUh9ZJAGq655S8UMKOJEr44owNnDRlt5zQQMueDVIdnxpemrCOMeOKFxQkyHwMkE8I4UwAx0GQTQQjkASAJgLvEGAdkzwZ5vj/27m2KjqcMcpPxmXG2iJzAXzSAbNEcCIIIwTiQcgxSQtA7xLTiyzy+1jg/THn+k11wxmjSg1NWMeIguLyccT0NQEWEfN4gCDiAAggXQ4kgEAAcfJ9kVcFdF/EtD841Ilr1x1jskb52ZcB8hXDNM03QOAA6wTSJUYiwBDBGEAESASyUwQPWQl+lL34rV1DGaNKLU1Yx4D8krKFxPw9InOCiMUBn/7DIQYRQ5xdT84uqa1Z89xQxNh8x+Rz/ZDcGfborMACcdv3GD0mhDwgHuAta+03Mxe/uXIoYlSppwnrA2zKlHnhvbn5dxLxVwUAxA34XMQGIq5NnNxQv6HyvkELEkDz8qJrQoa+bxgZsaAfyfQgviEQAYlA7t3JvGTqNVtigximSgOasD6gJsy9NKOl1T7IxlwsbpD6polAIFgX3NKwYc2tg3HK5jsnfTsS4qWBE9iB59P9iIBIiNAWd4/saXRfnLRsR/uRn1WlC051AGpIUEtr8MNBTVYAIAIRgWHvewUlZVcc6ema75z0tUiIlybs4CQrINnabYsJMkP8hdGj+O7BOatKF1rD+gAqKCm7goz3E3F2aAoggkBaIXxuffXKDQM5xb4Vkz8cZvkTgMhgJauuiICwR4jG3T9nL97+08EvQaWCJqwPmMLiRZPF2HUA5ferc72fiA3EBWuzMv3zdjz7YL+aXXLn+EiU/T+HfS5pTwxdjB4DTlAH4tmRa7ZsHbKC1LDRJuEHyfz5xrG9h8gMabICAHEWxP7slrbg5v5+bwv7/xoJDW2yApJTIsIeFThr75FHYYa0MDUsNGF9gBS+GbqK2Vw4ZE3Bg4izIOCmwpIvnN3X72lbXjQnxHTDUCerTtGEIBLiC1t2Tr5qWApUQ0qbhB8QuSUV0w3hORBGDnXtqitiA2eDlxHKOKf++Z8193as/PTkEW3NiefCHs04kukL/WUYEMG+IIF/GHHDtteGrWA16LSG9QEwZcq8MEPuJeZhTVZAspbFxj+d4rGlhzs22pxYlhka3mQFADbZNBxJBvduuntKeFgLV4NKE9YHwN5RuUvYeB8ZrqbgwcRZgOnreSVlH+vpmJblEy8wjKuHqyl4sGhCkBWmjxxn7ZKUBKAGhTYJj3L5JeWlxPQMgMzhrl0dgBgQ+XvcxM5uWlvV0PWtffecku8Fsb+GPD4pPsy1q65M8m5vs5Bzs67dXp2yQNSAaQ3rKDZmRkUWAfcRcWqTFQCIA7E5OWRDt3V9+eml8LwgfldmhklpsgIAK4DvUaYI7tt1x5islAajBkQT1lHMhtw3yXglqWoKHkxcABBdXjhz4fzO187FXBDwSizu9po0uNvaE4KsEJfkmKxbUh2L6r80uIXUQBSULJwL5sXpkqw6ERkW52YBgCwFv52z1c+8btsdgchCAuKUBp0Q7QmBMVjSdOfEuamORfWPJqyj0MjTLx0F2PsIHDpwMatUIzgbRGNxvl9+Mi6zbVTRr8f64Zda755cnH3t9t8mnFSHTOozlhPAYwp5TPfKT4pGpjoe1XeasI5CfijxHTL+NJH0ql0lh3Ak2vz2cbtBnAfgH/1sMxVOSjqO2MOpz1cAgFggiIT41NZWfDfVsai+S5PbR/VVwayFn4TD44CY9KpddSInkAsaqiv/1Lpi8iJiN6XdueWj9ma1tY1sfynk0Yf6szjfUEquXAqbCPB/spdsezLV8ajD04R1FBl7xmWFCT++loiLjmQxviFFDHHyUsD43L71q7Z3viwCar1r0jcyQ/zdaDw9EhYAhAwhbmWbZZ6dc82W2lTHo3qnTcKjSMKL3cHspW+yAgBxYKYzPCfP55eWPVAwc9EPC4rLFxBBsvZG7miLybMRP33+TsatIDNERcba21Mdizq89LlzVK8KSssvJuKHRdzR8TsjAhEna1w2sVec+3B9zZo3mpdPPtU38hyAUUOxDtZAEAEek8QCe8mIxW8+mup4VM+0hnUUyJtZPh7ACjma/sCIQJyF2ASIzSgi+hGKr/BHLN76t0Qg3/LTYLSwU8ecW/KYl7fePWV8isNRvdCElf6IRO4i9o5L66ZgL8RZkPHPz6eWawAg+4Tt/9ked09GQumTtBJWkOHT8c66u0SOoj8MxxhNWGmuoKT8Mmbv84O6NnsKiFiAeGlBySVn0sWwDHt1LO5qvTS6A6NxQcSnz7feNemyVMeiupdGt4s6WOHsiikg+oEcpTWrA4iAmbMBvm/87PmRyOK3tlmhGwxT2lRnBMlNWw3RD/beOXlKquNRh9KEla7mzvVc4O4h5oKUP9g8SDqahrOj1r8ZALIWb3uoPXCPplPTsHNZZZ/cPU8vhZfqeNSBNGGlqYLW477KxpuXbs8KHqnkssp8Y2Fx2RwChJy3uD3udqZTJ3zHssrzSkYV6bLKaSZ97hK13+jistMc83Mg5HxQalddJZdVti8jFD6n/vmfNbfeNenzHnOVdULpcrnJZZWliUFzwtfqssrpQmtYaWbKlKvDlnAvMX8gkxXQuayydzri0WUAkHXt9l8Ggfw8nSaUWgeEPM5JCO4TXVY5bWjCSjN7cxuuZ+Of80FrCh5MnAWRuTqvuOLjAGCdd1M04bakw2oOnZJrZ9E5rc5en+pYVFL63B0KecULZrIxT0OQ2fcHmwnp+RB0HxBDnNsUJz67uXplfcudk+f5vvzGOnjpUrlkAojQZgM5L2vJ9vWpjudYpzWsNDFmRkUWMyeXO+4uARGD2IDYA1HHnqACC7hEcn1iev99NsnnTdKdOLDxpobE3QYA2Uu2/j4WyH2D2TQkev9rf7E9vN9dqU4AnykTpMsqpwMdtk0TNuS+RewXHzBBlBhEDHGBg8hmca5GgJcgtDkO2tVs/WZYsiD4GV4sN0LBCUyYBlAxgDOYeSyIIOKQrv1h4gIQ05fzSst/11Bd+dgIY77VFrMfDfs8faDbgXUmJ2uBRECIJwiBJVhLEEkmLOo4jlngGcD3BL4v8IyA6MAfV3sgyAxzsYtlfgtAv3e6VoPnKPgz/MGXX1x+Lhn+H4gkVxAlBhFBnNsGkSon9FhDzTkvAFcmOr9HBIzf5WWjMR6Czw6lTW3eJLS/3/O1NGdEyZZzQ4QFBPoksckRsemZuIgBcW/BYXZdTeWuphUTPxJi/oMIQq4f4RIBzgHtMUY0RognGLa7rsAeWtHMgOcJImGHSFjgecnsJkg2DZkQjzv6xIjrtj4zsAtVR0oTVoqNPP3SUb6f+AuxmSYiIDYQF7wO8Io6co9gfWUTAMiPJowNLP9DAD5HRGaIyHgCckQoBECEEGWgjli2ElAdzpCnMq56c13MAph6+aT8kdGrCfRlYs5Jxw59Yg/OJtbUb1hdDkBaVhTdmxWiq9r6sHZWZ6JqjTJao4wgoP2v91dnPmcGIhkO2ZkOvicQAcIeIRbIxjjsnNzrduzt/9nVkdKElWIFxWW3k+ffABE4cS0kcoeV+N2NNVX7ACB6V9H5BPmygD4eMlTAnBxydwK4LrUlAiVrAZysDbQnBET4GwSPhLP3PUCX178bmn7pKTmR4DYic5Eg/ZqJRAwr7nMN1ZWPtd1VdCIELzEjt7dlaIiAaDujqYWRCJIdUYN1U4skf57ZmQ7ZmTaZxHxCS9zdPuK67TcNUjGqHzRhpVDh7IopzsqLbPxsZxMvinP/0lCzZj0A7PuPyWdnePItIczrWBUT/Vk/igB4huAbIBZInQh+kjFx2w/oIjTnl5b9C8H8gAg56fScIpGBc8HL2Vn+7B3PPtjesrzovqwwfaW7WlZnrWpfi0FblCEydOMMIkDIF4zKsYiEBE7QDEtnRZZs3TI0Jaqe6ChhCom1VxovI1ts4pcJE/9YQ82a9XLn+Ej07qLbwp485Xs0z0nyUZH+LnYnSC6Z0vFhL8jw6Zb4m0XroysmfqK+evWPrciFIvImmc5RxdTfCiIWzN7pLdH4PABgopWxhNiDExEREASEukYPrW28/7WhQgTEE8nyWqKMjBCNCMhdMXQlqp6k/i49RhVOm58tQhUuaH+orjVe1rS2qqHtrqITYyb0RIZHNwsQjiZkUFpt1gFtcQEbOoWZf9N2V9HNezesel6cvRDOVTnrHoW4NygNklayQ5wqACDC9ELgsMXrstUOUXLkr26vQTxBwzZ7o3PksGGfwb4WBhM+L3fM0GkOwywN7tBjk42EPwzgz3Uu63JsrIo3/ceEUwj4n7BP57XFBf0ZHeureCBwAj/i020tK4puq69Z80bt+pUX129Y9YW6XaPPck6WpzppJadgYM7I08py6ZotMRFZ27lmVmfNqn6vQRAMX7I6WF0jIxY3k6LhljNSE8GxSxNWqnhoipO5CjX3JxpXTJjoe+bXYY9O6cuo2JFwgo65SNIEAK0rJs4KfjT5a/KtX4fqJ8VuFGc3pDRpiYCIxnCGnAIAzPRC58ROjwh7m1KbrIBkc7s96lFTM5cc9mA1qDRhpUjDulXrmqtX1suPpmWHwJUZPk2NJoZ+1I4IaEuIE6LfyNJpIRF+1GSZe9pacROqqixAz6a6P4vYEAtPBgCIbLcu2VJsjmJFNE57TRrctc4RmprNSamO41iTBr/6Y1trLHprZpjPHuqa1X4CGAIZh5HAxgAkNbbN7nWQv3UcUDB4zyYe+LhQ8rGivtxyBJCM6ThDgzEE61CVf/PWxQT5fjgN1lUWEQiSMarho4/mpFDz8qKPegZXtQ/jxqICIOwRtSfkKlqGv8ijkUva3k4Ujliy9R1Mmx8SorMGpbVFDDgXOGf/AJJnyKIeRMcL4QIimpMMpufrJqHkki4Mao+7nSTmOgCIh70ViAfzIh6f2z7AR3cGgwAgQihlARyjNGGliCydFmpB+62+IdM+DE3BrtoTgrBHZS3LJ+2hizdeB+AdAMDGqjiVlH1PRNbgCOboETEcZBMEV9ZvqHzmoLf/vaC0/GIQ7gNxfo87AYlEAcCK9QDv+hFLNr8DANOXbYxv+tcpX41b+ovh1O1tSAAYFE1N6ccuTVgp0pYb/XTE8KzhTladiABibD349boNqx/JLymfx8b74oB26iGGiLxLSHy6ruaRTXnF5RewoSshmAKgQUSq6qor78svLqsllidAFDm0piUAyS4AyA5GVuOGV9q6vjv1ti0bN//r1G9mGP6RHYrh1D4gIgiSMarhk/rOgGOUEzohVZ3HmSFCW9w9dnvj9vu6e98PQjc6F2wdSOd78qFt+fe66kc25ZeUX8WGnyQynwPRDGI+lz3v3vySsqr6mtXPEOSm7vq0xFkH4U0AQDe80krddKpN2XL6j9sT8kSGn8pb2G1MYeHHJE1YKcJGnm9PDP++874htCdkl3F8zbJl6LZB9d5LP68lkauT6231AxGcDeqDiDw6uqSsCITbARhxASCuYyfoAOz5/1RQUvbPtdWV94gNKom7VPSJAchOCYc391pUVZV1cF9PBG5P14mlw4EIaE+4wBGqh7VgpQkrVeoTiVetw2ZvGJcE7iwpsFicef3Wt3s7tm7D6t+J2PsOSCaHPT8BhB37nlvdKESfYvayuu2jSjYBLwaARAauFhe8QpxclJCIAcHT9c//rPlw5Z38/c3bnHM38iA+8NwXHhMc5PWE72sNa5hpwkqRE5fsjBLJw74ZvjIjIULMysoRS7Y93JfjnSS+JTb42/4VTvvBAfk9vikCgEYBwL7nVjfC4TIRty/Z/+VEiFb2tZzJ39/8UCyQR4ezaegzgUTWTF+2MT5shSoAmrBSiiw/EE1Iw3D0ZYUMIRqX7dlh9HlDheQSN+6rAhfvSx1GICDBhLyZ5TkCvNzjfK5ks++Nzv+tq6l8gcRdS2wg4p6vz9z5bF9jJEDCYVzXnnA7h6O2apjQFne1CTa/GPLC1CE0YaVQ5vVb33bWrQh7Q/tBIwIcxDmSr9NV2/b053vrNqx5VkTu7Gyy9UoEZLx8EilrcFm/dTZ4+ZAmJTHE2Zglubfry7XVq38hQfweBn0Xzz7br+HJCcv+/q5zuBYCGepHdjp29bl92q1v6AhhCuh6WCkmPxmXGW2L/CnDp9lD9WhOZpjQ3C4/ylm87eqBfP+44isyE9zyFLE367CrlRJBgPc84Gwr5INcJZEp6bzVxLk9Im5J/YbKVQOJpTebbp56f1bY/HNbfGgmZ0V8RjRh/+yFYp+YtGxH+5AUonqlNawUoyt3tQXWXhYP5N2hqGklpzDIMyNs9oA3T9hVc3+bOPmiOPfuYR+tEQETjw2ce7DORbZncPwjztrPOJEb4OxlVuIlQ5GsAABhd300bp+PDEF/VnJ5ZPeWsPuSJqvU0RpWmmhbMfnDhuVXnqGxgzWZNDNMaI/LWhu3n82+ccd7R3q+wpKKsx3hV8w85nA1reQa7fEf1m9Yc82Rltsff//GScd7ML/O8Ki4LTE4Na0Mj5Bw8o4TXDTl1r9vGJSTqgHRGlaayLxu618Tzl4YD+TVzHByffaB8jhZs2qPy68TgffpwUhWAFC7YdXz5OSTIvJasm+q5yDFBWD2rs4vKa8YjLL76uRbN78D8T7VHrgnIz7jSAY0mIDMECNu5UXnaJ4mq9TThJVGsq/b8VLchM6LxeU/CUhEQv1LXJ2JSgT10RhuyGjc9k8512+qG8wY62oqX/Di7R8VF9wPQqK3xCUQIqIfjp51yYzBjOFwJt/22u6obz4bS9h/BdCY6TNMP36QTMn+KiaKtSfknlYXOn/KbW+8NnQRq77SJmGaavuPyWcbX651ggszfMoWSS51bA/eKYeTicoJEA/kPZA8Aph7ItdsOeQ5wcFWWFw2R5ivFWAes8kGBCKCrtMZiA0kCP5GFP9YbXXVoNT0+mPjzadMDbF8nQn/5DONISJYJ7CCjliTjxMxAYaS/44FthmgJwLn7jr5+5vXDXfMqmeasNJc+11FJwnJBRCaC+AUERQKEAHgCNJKoF0CvALgT5bkqRHXbt893DEWzFpwklj+OJjmksiHRFBIhAwkM1grsVdrXfDdhg2rfzXcsXXavnTaWBfY8yFyPgSnOdA4EWSBQASJAqglotdJ5BnL7n+nfm/oE77qP01YRxF5eq6HF7bmIJSV0ZqISjzuteXetK0pOZMgTcxd6uXENub4MZNBbMQi3tZYU9WEwVsV8IiJgLbdXJRjTCiTfCGCaT8BhU20rH/zv5RSSimllFJKKTVstA9L9aqgeMEnhb2ZEOuB2EKkpn5D5W+QRn1SA5VfekkJ4P8jxHoACQAwG7E2/v8aah7+Q6rjU4fShDVMli4Ff/u4cRldX6Mrd7V1d+y44isy+3LOILuRXWJEl99hHdgPDyiR1D5b1XLwawXF5beQZ757yG3iErfWVq++ZSDlpIsxMyqygrB7iU14ygFrdhHB2USLDWT63hdX70hdhKo7mrCGQWz55FPFyH9BMDZwAAjwCbCQFyMhezld9VYj0PGQMbXcBzbniLOH/d0QoZslFKifCUsIYAGwzTF9tWHdytcBYOQ/lOV67XiD2Iw+8APNgLh9BO+U2uoHh31e1WAZM6NidBCS14kp75A15QVCJGfVVq9+KSXBqR7pJhTDICC3KDPTfDiICkyXFBMK8aS2NnkQwOMAEDOtMwz7l4oTcA/LuQxVO4yNN0GC2HcA/BMAUCtyYJB5SIkiEJGIoG0kgKM2YRH7AupxWQcH4qO+yftBpAlrGLDj38ejbqG1kte5yYtnSBJtsomsebHzuAzhzTEb/IVAxa6H5d4J5IHZ9Lg91gAlH2amcfvLYSME29OHVkg/0CoFNGENg8iSrU/vuqPojNwMjGJLEguEshhSF/ffK7zp7/vXLn+3emX9+NnzP94u3nhyjsCHJgVyyBLnlhLzRTLISUupdKcJa5iMu2HbHgCHXe1z59qqKIBed4zJK15wG4Mu6vVERCD0fQo8sQE5q4+jdCB2+tcgDWnCGk7FV/hHeooxibaQZfk8gXvZhYsg4vYIsLNvZ2XYIL7JSTCoI39jZlRkuUyMEyfjxCGXQb4D4iyugQ2/u9tm7kTN/YmBnj+3+IqRoLYTmGSMADkQJkPSZp1tEArvamxr3Y2NVf3fKCK5dXXW++XMH2n80GRYOc6JCQFoYsJbdZlvbevvcs7qyOgo4TAoKF04FST3QeS4Q3c57hcRogwCFfXW/U7swbnEXfXVqxf3vZv+wNHF3OKFJxqyr4F5xIExE0RcjMjOqKt+ZNPBZ8mbveh4tvZCAPMAnCmCccQUSW4BRp2d9hBIC4G2QfAHFvfgnprVr/YlyglzL81obU18BuAvgKQEwFgQ+9SxmLuIAOIAkSYQ7xTQX4mDH9StW3NArXXsGZcVJvz4RiIqOOR3QgyIfd6B7mfBbBBdCMgJRMwgSsYvLkqgjQR50HP7HthV89tup6iowaU1rOEg9kvsRc4X26fNZ/pwvj4kIWKH5J7vg1Dg4Y09Y35hwve/QdaVE3uFnUvNECSZpCAH5E4CZXfsBj3DCf1LQUnZPaP2Nnx7y5bfx3oqo7C4bE5r1K4g45UCgIhLnlPcoT8S5hyApjGbaWLdnJzZ8+c0ra1q6NPFiAPInG2Izk6W03ENneUl448QUTHIFCcw6gtjZl1asXvdg2/2/SemBkIX8BsGRPK0s7F6EbHi3BF9wTkLEZtsD4pLbpnVTVISGdbac+D5JWxC14JQKC5IjjpKR87slnTsBh0AkCzyQjc35ub9vKdmc2FpxTxh/h0Rl4qzHec/cO2tA0/fcX6bALE5JeRCH+7XBXXuVL3/Og6NXzriJzZzAhf8Kmf2/Lx+laH6TWtYw6C2es3/5JZ84UzfhHOBwdt7U8QjgS0BeDlAOal8WsaX7GfjQetLzHyG9DcOEYhNgNlfkE8tL9UDt3d9O//MS44TyAMgHtHtWvLEnbtOv7+AYNcqV/LfTf2+qL6G7wKw8c8MBfgG0Pd9H1X/acIaJo0bHnkbQK/bww/QywUlCz5Nxv9M1w8zQYa1M3hXzf1t+aUV/wGiVRAkk0hyw9SOWgosiAwxv9/PdBARCwjdUFh66UNdZ9GL4UuZveOStbGuOvut3DaBvAdHAYAcAOOIaEznXorOBQ/Xu6y1A7qwjtHWZN+V67E53hHbZQXF5XfW1VTqnoVDRBPWMBN5vzLQ4/sdelmYb/8xo0vKJlmiUw+sUTgIcHpeafmlBr01DRlWaGfDp4qewrJlRzyMn51pftnalriZTHi62Phb4twTgPuzE2wTUNQzPNK5YBZA/0LEUw5JWiJg9grEBp8E8LP9URLOPfTHRRBIKxNfZtrx5O5XKluTry/lccWb8hJE0wQ4w1na2dAW/y02VvZ7NJLYQGzQLMDfQBQFMIWIT+h2/psIyJg8EXsugDX9LUv1jY4SDpPW5ZM/E/JxQyJwEd+wiwfybKbhW+iaLTEAaFteNMcYfNeK5HTmnpBhG7PyeNbx226ji2Hzi8vPZUPLnCAb6PzQ0Indj3QRDruHIAgdNaD76ybGr0JV1f4q2kBHCfNnlp1H4DNiQg81V6+s767UvJnl41nojyA6+eCkRexBXLCyrrpyUedrBaVl64lMqRz0kLI42U2CMwdSo+l1lBAdycq5xwG6oa565SYAyJk9P88P/G8zm6u7S1rJ0dlgRX115eL+xqP6RmtYw0CWgtvIfdfzzXQCgwjIDFNJW4D/BrAWAITkm6EMc66Nv//hYQaMQ2n725MfBrZuIZalZEIfYWs7ht47R9+6qYiJJJtYh0UA5PK8HaG7G4CNR3qt9etXPw3g6d6OaVhfuTO/pPwBZnP7IR/85DVNPijGvckpEQceR0xjIPhrQWnZ7yH8gjBvZnZvoyn2Xu3GQ1ef6DNiiLOvUWu8vOt5mtZWNaD4iiUF0nouiE87tFkrADBhwOWqw9KENRyWQWQF/hQk5EOBFWOY0B641yXk3uw8RBw9FcTc+YGD3/m59ITgRJ6Petw5Q/5PzgVzIOIf4XyuLgQAsziMGqQTYvzs+ZFoEJpOTDNEUERw+SIU7nyfkv+Z3V1CFQgIyMHcud7+SZkifwHxx3HwRFkRgGgCkX9l56ijCyRKmaH38ksrNhLkKTL0eO3aVVv6Ez8RQ5x7oNukV3N/AiXl/0vMp3XfNHx/wqkafJqwhgEBIsdvX9L61pQHyaewYZJM399CX3tjf5Mpe8m2O2J3Ff3OkM2OBh2/GGaXyeaNrGu2NAFAXfXq744uLvt14HEmOwlDUAk244/sQWgC4ISY9h7RRSKZqNpt6Cvtjr5EJB8iMpycz2kOeUiop4735JswaDmZgGcBAFbMz9kmvsJsDt1xWgTSZXyBiCIgmsREkwD6lDj7bwWl5f9lYvTt3a+sau3LdYhYkOHeNk3t8QkC6rXPUB0pTVjDhC6GBba82Nsx4Wu3HXazzq4zwvNLytp6/HQQo3P2dy9RARA46x5omBj/O9YfrvSejSteUNBu+RFi76PSMf+qp1G9pL7XEBtrVr6VV1JW4UQeYvbGJUfrekrSnZM8O4ukkWT8620oPhnT5l/Sp0d1nHNgbu7pbaIen4lSQ0wTVirNn2/yt4euBPE573ei95HAAOj2UR8ihoj7g0AegfSStYiEhHbWT0r8sWuH+wBQgmkFGf+jYg8djEt2YFsIJNa5eykRZRxyYC8aNqz+45iZFXOs2OsguAhEJxB1rhnWZQPXQxbjS87xIuN/Ni8LZQ3AL/pwObp0TprShJVCedu8YvLMvQTCQAZsk82jbj5byWfhXqpbv/qBPp2out9FHyBvZvmHILhYbLdTv2Li7HIL9xsWbiDTUTWyMh/sfa8/zdnd61dtB/D1kaeVLfXDZoYgOBOEDwGYBOAEiBxH7OVItzPsBSRYgD4lLJWuNGGlELO3Fy7YC/ZHDegEHQ8Td0dAw/e7ta6EPD90cP9Sch6TXVO3ofIbB39L3qyK6oEuAbjv1dWNSHZwPdv52pQp88J1I/LGehx8DODbAcrrmrQ6KnYTp02bH9o4kBUcVFrQhJVCddUrN+UWl33EI3eGhdAho2A9sQAxFRJoMYiOdAWII8aE3O5riASQ1Hb7PdZ9lozX99CXLuX8J7ecR8IF5ElN7dpVW9ElI3U8NL0DwAP5JeVfY+Y86ebk8XizdoofxTRhpVhjshO9T0urHCy/uLyFPf5x3+ZbDR0hqu1uUn4yLl4wunTBY3uq16wFICNKF+aHIV8BcEW3zwV2p/gKv+CJzT8l4y0CAGdta0FpxasgeUlEtjBQb0FCgtFE8jEAMw6eckDJxwa29bYahEp/mrCOZoTxqQ4BAIil2tmgnYgzDug7EgGIxzvhpwtKyl4VoJ3gJhGb4/ucrAAUmtbzQN4iccm+KQJnEdFskJmdHFKQju2D9s/c7y5KgHjVEVymSgOasNKAPArT/HbR5REfJQkLMkaaowm+d9SSrVsAIK+k7HPEPA8i/P56TJIPwif788Hvd1w9T4wg6bJJRt26NZsLSssfJeMtOmSUUBwAChNzCXU8QNzfmJ1IEXMyGXVG1us8roODNT7Exh+vb01U9atglXY0YaWB9ncmnzMiAz8GAM8A8BmBdScB+MeRMysmscgaYi90YIdPTzWJQYtqH1GoCaDsA1feI0DQ4gfZByyGx2xvdBankvGLk6OFB8Uq7y/gR2QgzloQmb7Mx3LCz5G1UWIT6axl9UXnag1ig/9OtOOKrnOwxCUI6G5fx+TloIddiwBABKaXjrCezqkGgS7glwacSEt7QmLorM8kP5O7AcC30i4i3ayUSSD2evwSIHQkMTXWVO0jkVsBJLqeF8ku/++/99LPD+hM37Pu4d0mRp+EDX4GIJo83hz05QFgKy6ohJFZAtlKxu8udg/Z4/ZnpcYNq15z4ipE3AaBtHee69Dzm/fjFAQi7q/OSXld9eQvdIws7sd+ZguAncQHlW98ANJgjdftYAEACGgrIN383A0A6nUDEXVkdMQkTTQtL5qTYTDDipCQNLdH+fG8m7ftA4DckorpzDwHzhpQx6RGEWKCkS6PghBILJEYIQLjxdp1K5870rjyiitmEeMshnhO2LLYl2trVv+lt+/JLb3kVEPeBRCcBcgYJGvyjQJ6zcD+fk/1mr8CQEHJgrlE3gxL71cdjRABsrm2etXvDznx/Pkmf2fmSZRInC7gU4lkogAFBGQk12OmJgDvCPg1hl1bW514Feh5QmzBrAUnkTMfsxAPRAIRMkTOSbCuvvrhXh7NWcr5JZs+xeRN7IzdiBCEWuIJ89i+lx/c2+sPVSmllFJKKaWUUkoppZRSSimllFJKKaWOHYT583XmehrRR3OGjVBBafnlibhfNRwTC8fMqBhtQ7IAhnMQBDV1NVN/D3S/92B+6cIScsHoupo1Tx7uvCNKF+b3tH1XXxTMLD+HErSr9sXkxhC5JRXTjUjIguvISDkBCXKuCR5tgMWUuurKR3s6V96shR8Slzitsfrhg4+hgpKyeWDvTIjdk5Xprdrx7IPtXQ/ILy6/sj1Bj7W+smoPepBbUnGq2W6n1wEPD/R61eDShDVM8mZWzBKHhZ6faAewsvDMiinwaDJYxomj5+uqi7YUlGz+BBlvrBX7h4aW+J78SOhssBxHRPvq1lc+UTiz4nTnhbbV722OFURC0y3iW4n8UkN0PJi3ds5sL5w2PzsIyQ9g8JgAmwk0dfzsjeG2oGyOMd4JCGJ/rq15ZGtu6SWneuyVinMngHnXtGnzQ3sy/c+CSIj53dp1K5/LLamY7nlULDb+1wT5CR+4L1y84B72pCYIvBnGw3g492Igtj7Evrd7/artBbMWnOSEfXZuBkCS4SUe37m2KgoA4nABPDkNwEVYupTNE1vuFpYqcrKVBBlG3CNkrLNBaIKQKxo965IxVvwLRFxLw8T443k7QuOM8PlW7DuANBsx5xWWlkcgtLl2w6rnASC/eMHJIJRZx98wiBf47+TJmFmXTnSQ81xgt9fXVD4DwjTj8x/zSy85sb764Q0FxeXjyFGWsLQS8wXWybuGzKuAe2PMjIrRCV/OMB4dHzhebzgRl7i01r/48Lu5pZecSmSEHM4SQmuDy/otau7v96atqm/0WcJhwk4uIrKXg1CM4it857nPOLj5TniHwN48ZubWCQL24MSxwy15OZFCItwqgj0AlRSUll/snPs84rHjRub5WcK0yDPhiQz6hiN6W5y7omDWgpMAAJnhGYC8Xb+u8vGGdStfF9/bEGuCYUKGgKKOzS0FsxacxOAlDuZNACcRJNid6X8d4JFgbhHn/q2gZOGZDHeNg2kQmJtY7FiIa2Xn7xQxpxO5a5xghwiuY/EmW3FfwdKlDMtXknCmMEWF+bSo9b7Y+XMgoE6A9vzisvMKntj8CSHZA6CJ4AIIxokxZyQCHm2MtEMkIPEyk4kJ5+bvCH+andwoxB4MN5MjC0iRI3pbyH1ldElZEQAEEdoN0D5j7BcJXqQtd292YINvOuEdRLgwr7ji44A0e8aGRfhyALDCM5yRC4TlJgH5MNIcID7JCX3cZcgZTPIV58xbRuz1YnEKefwlFF/hM8yX2SIDzM1GcG6hablouO+tY4kmrGGQN7N8mhDOFvE+w8RTC7ntfACtJPLL+vUPPU2g7VaklERmgEgEkmfgIiC80LBh9R/F2kdEcBoIwmICjiYcxLE4YpA8Vb9+1VMi9AocjQMAy7KXCGM6y6cgWOyyMk4BcAYnn5vLFssfAuGN+nUPPkNCvxZBiIAJcS/233XrVj4JkW0QdwoIo2BjOQBvMvDqRfDOnhcfeqXjzE/Wr1/1lBB2sk87AWoueHLzlwHZxC7IgOOTiEwUkowLAAQIichPQbQAoPMJqGThMIwjIYrDodUYExMRIiJnIdNY5AQh0w7BOCeyimBHs5MLROwoISSvH/S3BHgsAOx7bnVjXXXl1cJcJXCL4mQXgbCzfv0vngLjt0Q4vWMnVun8ABgKDENInKwmcoXscAGDM4kpgDgC47f11Q/+CSS1BvQ3J8jK55YvU3Ij3HEs7kQh026Fxg7PXXVs0oQ1DFjkcwAtM8Y+5IxZLJDzSCgDwD8Xli68WiD5ztEOEMZbsUSgIEhQHIKzCkorvk7E1zjQ7wT0kiP7debwjQBlsNiAkNyUjwSWvOQuEA3rK1+H0Ov5JeV3F5QuXAzQSHLOidCJ1jmBCDPF15NgQkf5C4XQQuA1fuDfUlBa8U0QJlnynwPwjjGhEQSpr5vYvhUACkrKvyQioc6yAUmYmEuI4FER+lqMzC8d0SQiCYtLhIiwf1kXIlh2aATx/Zbsz8WhFdKxNrRITW31yv+pXb/q5eQ1SUwcJoiABUGIxFlD5gQhUydOJjJTmEDJ5leX688tXnhifknZd8i5c0CwsPQsCTIKZi26Bg5ljoLfAQyKJ/YJeE9h6aKbhKnCCQX7z08ygZwLk7UJErIkbAGAgAT53MJMT5Lgyxkm/htApu6P0UHXix9C2oc1DAj+T+uqH3yv4393j5lRcWcQcotIqEo89/cIEj/dWV0VHTOzYo91yPIC/0n2YiMc8wZr3dPi8Ou9L1buAIDc0ks2GaZ4nIL63Lrm1tqxeTsBIBGYn4/KNp0dy1K3ofKuwjMrpjjjchEK/1fd8z9rzi35wnd8plwAf9yzrmr3mBkVN9pMd7LE7C85xzahzZ9CzNXiYAg0tnHDL3aOmVFxi43gFDi8i6oq64rn/5vvmQnWZG6XIEYAkDCZPyRuyyeYM8TZ/5vslJdVo4vLp4vjaMCZ+5dqScS9n42ItMV2rl0dBYAxMyreTGRITSbbRKu1NZ3HRSLeS3v34o19U9uaC9/0TzPEj7vWWK0JZUcSvkwMs/nVuzbSNDLS8jwABDH56aisULQBQGPNyrfzZpb/WIiO89g9vqf64d3Tps1/bU9O1nQieqSxuvK9ccUL7t41xTZiW853CrzW01xgV3tR12hCiCT8rImhQB7bhRH7RgaxrEyvPh7jkR4A+DDL260bw4KTQbRm59qqKObOvaew7fjphvjxIByqG9q76dimy8ukSH7xglOMJ4171j28u7v3C+fOz5aW0NS6msoXhiumvNmLjifrzmWQJ879b11N5a6+fm9+SfnJxDzH2vZfNtZU7RvKOFOtoLj8LGKcFjbxRzsHE5RSSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSSimllFJKKaWUUkoppZRSqfH/AXlTRJE7lZLQAAAAAElFTkSuQmCC"


def _get_secret(key, default=None):
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default


_GH_TOKEN = _get_secret("GITHUB_TOKEN") or os.environ.get("GITHUB_TOKEN")
_GH_REPO = _get_secret("GITHUB_REPO") or os.environ.get("GITHUB_REPO")
_GH_BRANCH = _get_secret("GITHUB_BRANCH") or os.environ.get("GITHUB_BRANCH") or "main"
_GH_ENABLED = bool(_GH_TOKEN and _GH_REPO)


def _gh_config():
    return _GH_TOKEN, _GH_REPO, _GH_BRANCH


def _gh_headers():
    return {"Authorization": f"Bearer {_GH_TOKEN}", "Accept": "application/vnd.github+json"}


def github_upload_file(path: str, content_bytes: bytes, message: str = None):
    if not _GH_ENABLED:
        return False
    token, repo, branch = _gh_config()
    url = f"https://api.github.com/repos/{repo}/contents/{path}"
    try:
        r = requests.get(url, headers=_gh_headers(), params={"ref": branch}, timeout=15)
        sha = r.json().get("sha") if r.status_code == 200 else None
    except Exception:
        sha = None
    payload = {
        "message": message or f"auto-update {path}",
        "content": base64.b64encode(content_bytes).decode("utf-8"),
        "branch": branch,
    }
    if sha:
        payload["sha"] = sha
    try:
        r = requests.put(url, headers=_gh_headers(), json=payload, timeout=20)
        return r.status_code in (200, 201)
    except Exception:
        return False


def github_delete_file(path: str, message: str = None):
    if not _GH_ENABLED:
        return False
    token, repo, branch = _gh_config()
    url = f"https://api.github.com/repos/{repo}/contents/{path}"
    try:
        r = requests.get(url, headers=_gh_headers(), params={"ref": branch}, timeout=15)
        if r.status_code != 200:
            return False
        sha = r.json().get("sha")
    except Exception:
        return False
    try:
        r = requests.delete(url, headers=_gh_headers(),
                             json={"message": message or f"auto-delete {path}", "sha": sha, "branch": branch},
                             timeout=15)
        return r.status_code == 200
    except Exception:
        return False


def github_download_file(path: str, local_path: str):
    token, repo, branch = _gh_config()
    url = f"https://api.github.com/repos/{repo}/contents/{path}"
    try:
        r = requests.get(url, headers=_gh_headers(), params={"ref": branch}, timeout=15)
        if r.status_code == 200:
            content = base64.b64decode(r.json()["content"])
            os.makedirs(os.path.dirname(local_path), exist_ok=True)
            with open(local_path, "wb") as f:
                f.write(content)
            return True
    except Exception:
        pass
    return False


def github_list_dir(path: str):
    token, repo, branch = _gh_config()
    url = f"https://api.github.com/repos/{repo}/contents/{path}"
    try:
        r = requests.get(url, headers=_gh_headers(), params={"ref": branch}, timeout=15)
        if r.status_code == 200:
            return [item["name"] for item in r.json() if item["type"] == "file"]
    except Exception:
        pass
    return []


def _gh_check_connection():
    """Tes koneksi GitHub SATU KALI (dipanggil sekali per sesi lewat
    st.session_state) untuk memastikan backup otomatis benar-benar aktif dan
    bisa diakses - bukan cuma cek token/repo terisi di secrets. Ini penting
    karena kalau backup GitHub gagal/tidak aktif, SEMUA data yang di-upload
    user (Target, Project Tracker, Omset, dst) akan HILANG setiap kali
    aplikasi Streamlit Cloud restart/redeploy/bangun dari sleep, karena
    filesystem Streamlit Cloud sendiri tidak permanen.
    Return: (status: bool, pesan: str)."""
    if not _GH_ENABLED:
        return False, ("Belum dikonfigurasi. Tanpa GITHUB_TOKEN & GITHUB_REPO di Secrets Streamlit Cloud, "
                        "semua data (Target, Project Tracker, Omset, dst) akan HILANG setiap kali aplikasi "
                        "restart/redeploy/bangun dari sleep - ini kemungkinan besar penyebab Target dan "
                        "Project Tracker kamu hilang berulang kali.")
    token, repo, branch = _gh_config()
    url = f"https://api.github.com/repos/{repo}"
    try:
        r = requests.get(url, headers=_gh_headers(), timeout=10)
        if r.status_code == 200:
            return True, f"Aktif — tersambung ke repo `{repo}` (branch `{branch}`). Data otomatis di-backup & dipulihkan tiap restart."
        elif r.status_code == 404:
            return False, f"Repo `{repo}` tidak ditemukan atau token tidak punya akses ke repo ini (HTTP 404). Cek lagi nama repo & permission token GitHub-nya."
        elif r.status_code == 401:
            return False, "Token GitHub tidak valid/kadaluarsa (HTTP 401). Buat token baru dan update di Secrets Streamlit Cloud."
        else:
            return False, f"Gagal tersambung ke GitHub (HTTP {r.status_code}). Backup otomatis kemungkinan tidak berjalan."
    except Exception as e:
        return False, f"Gagal tersambung ke GitHub: {e}"


_SYNC_DIR_PAIRS = [
    ("data/main", "data/main"), ("data/ads", "data/ads"), ("data/walkin", "data/walkin"),
    ("data/target", "data/target"), ("data/corp", "data/corp"), ("data/log", "data/log"),
    ("data/projects", "data/projects"), ("data/_cache", "data/_cache"),
]


def sync_data_from_github():
    """Restore file dari GitHub ke lokal (dipanggil di awal, setiap sesi)."""
    if not _GH_ENABLED:
        return
    for remote_dir, local_dir in _SYNC_DIR_PAIRS:
        for fname in github_list_dir(remote_dir):
            local_path = os.path.join(local_dir, fname)
            if not os.path.exists(local_path):
                github_download_file(f"{remote_dir}/{fname}", local_path)


def backfill_local_data_to_github():
    """Upload file yang SUDAH ada di lokal tapi BELUM ada di GitHub (arah
    sebaliknya dari sync_data_from_github). Ini menutup celah nyata: kalau
    user upload data SEBELUM backup GitHub diaktifkan (mis. GITHUB_TOKEN baru
    diisi belakangan), file itu cuma tersimpan lokal dan TIDAK PERNAH ke-
    backup - begitu Streamlit Cloud restart (mis. gara-gara Secrets baru saja
    disave), file lokal itu hilang total karena tidak pernah sempat masuk
    GitHub. Fungsi ini menutup celah itu supaya begitu backup GitHub aktif,
    SEMUA file yang sudah ada lokal langsung ikut ter-backup, bukan cuma
    upload baru sesudahnya."""
    if not _GH_ENABLED:
        return
    for remote_dir, local_dir in _SYNC_DIR_PAIRS:
        if not os.path.isdir(local_dir):
            continue
        remote_files = set(github_list_dir(remote_dir))
        for fname in os.listdir(local_dir):
            fpath = os.path.join(local_dir, fname)
            if not os.path.isfile(fpath):
                continue
            if fname in remote_files:
                continue
            try:
                with open(fpath, "rb") as f:
                    github_upload_file(f"{remote_dir}/{fname}", f.read(), "backfill: file lokal belum ada di GitHub")
            except Exception:
                pass


# ========================= Cache helpers =========================

def _dir_signature(dir_path: str) -> str:
    """Signature ringan dari isi folder (nama file + waktu ubah + ukuran), dipakai
    untuk mendeteksi apakah data benar-benar berubah (untuk gating log/backup,
    dan untuk validasi cache parquet di disk), tanpa perlu baca isi file."""
    if not os.path.isdir(dir_path):
        return ""
    parts = []
    for fname in sorted(os.listdir(dir_path)):
        fpath = os.path.join(dir_path, fname)
        try:
            st_ = os.stat(fpath)
            parts.append(f"{fname}:{st_.st_mtime}:{st_.st_size}")
        except OSError:
            continue
    return "|".join(parts)


CACHE_DATA_DIR = os.path.join("data", "_cache")

# Naikkan angka ini setiap kali logika pengisian kolom di load_all_main_data()
# berubah (mis. classify_pilar_hybrid), supaya cache parquet lama di disk/
# GitHub otomatis dianggap usang dan di-parse ulang dari Excel - bukan cuma
# dipakai apa adanya walau isinya sudah tidak sesuai app.py yang baru.
MAIN_DATA_SCHEMA_VERSION = 3


def _cache_paths(name: str):
    base = os.path.join(CACHE_DATA_DIR, name)
    return base + ".parquet", base + ".sig"


def _load_cached_combined(dir_path: str, cache_name: str, required_cols=None, schema_version: int = 1):
    """Coba muat DataFrame gabungan dari cache parquet di disk, dipakai supaya
    cold-start (aplikasi baru di-deploy ulang atau bangun dari 'sleep' di
    Streamlit Cloud) tidak perlu mem-parse ulang SEMUA file Excel dari nol tiap
    kali - cukup baca file parquet yang jauh lebih cepat, selama isi folder
    sumber (dideteksi lewat _dir_signature) belum berubah sejak cache terakhir
    dibuat. Beda dengan st.cache_data (yang hilang tiap kali proses Streamlit
    restart), cache ini disimpan di disk (dan di-backup ke GitHub kalau aktif)
    supaya tetap ada walau aplikasi baru saja restart/cold-start.

    required_cols: kalau diisi, cache yang skema-nya sudah usang (mis. setelah
    update app.py menambah kolom baru seperti PilarExcel) otomatis dianggap
    tidak valid dan di-skip.
    schema_version: dibandingkan dengan versi yang disimpan di file .sig - kalau
    beda (mis. setelah logika klasifikasi di app.py berubah tanpa menambah/
    menghapus kolom, seperti classify_pilar_hybrid), cache lama otomatis
    dianggap usang walau nama kolomnya masih sama persis. WAJIB dinaikkan
    setiap kali logika pengisian salah satu kolom cache berubah, supaya
    perbaikan tidak "hilang" gara-gara cache lama yang masih dipakai."""
    parquet_path, sig_path = _cache_paths(cache_name)
    if not (os.path.exists(parquet_path) and os.path.exists(sig_path)):
        return None
    try:
        with open(sig_path, "r") as f:
            saved_sig = f.read().strip()
    except Exception:
        return None
    current_sig = _dir_signature(dir_path)
    expected_sig = f"v{schema_version}::{current_sig}"
    if not current_sig or saved_sig != expected_sig:
        return None
    try:
        df_cached = pd.read_parquet(parquet_path, engine="pyarrow")
    except Exception:
        return None
    if required_cols and not all(c in df_cached.columns for c in required_cols):
        return None
    return df_cached


def _save_cached_combined(dir_path: str, cache_name: str, df: pd.DataFrame, schema_version: int = 1):
    """Simpan DataFrame gabungan ke cache parquet di disk (+ backup ke GitHub
    kalau aktif) supaya cold-start berikutnya bisa langsung pakai cache ini
    selama file Excel sumber belum berubah. Kalau gagal (mis. ada kolom
    bertipe campuran yang tidak bisa diserialisasi ke parquet), gagal diam-diam
    saja tanpa mengganggu data yang sudah berhasil dimuat."""
    if df is None or df.empty:
        return
    parquet_path, sig_path = _cache_paths(cache_name)
    try:
        os.makedirs(CACHE_DATA_DIR, exist_ok=True)
        df.to_parquet(parquet_path, engine="pyarrow", index=False)
        sig = f"v{schema_version}::{_dir_signature(dir_path)}"
        with open(sig_path, "w") as f:
            f.write(sig)
        if _GH_ENABLED:
            try:
                github_upload_file(f"data/_cache/{os.path.basename(parquet_path)}", open(parquet_path, "rb").read())
                github_upload_file(f"data/_cache/{os.path.basename(sig_path)}", open(sig_path, "rb").read())
            except Exception:
                pass
    except Exception:
        pass

# ========================= Constants =========================

MAIN_SHEET_NAME = "Faktur Penjualan"
_FAKTUR_SHEET_NAME = "Rincian Faktur Penjualan"

DATA_DIR = "data"
MAIN_DATA_DIR = os.path.join(DATA_DIR, "main")
ADS_DATA_DIR = os.path.join(DATA_DIR, "ads")
WALKIN_DATA_DIR = os.path.join(DATA_DIR, "walkin")
TARGET_DATA_DIR = os.path.join(DATA_DIR, "target")
CORP_DATA_DIR = os.path.join(DATA_DIR, "corp")
LOG_DIR = os.path.join(DATA_DIR, "log")
PROJECTS_DATA_DIR = os.path.join(DATA_DIR, "projects")

# Catatan: cabang ke-13 adalah TELUKJ (bukan KARAWANG - ini sempat salah di
# versi sebelumnya, dikoreksi setelah membandingkan dengan file Target Omset
# asli yang di-upload user, yang berisi 18 cabang persis dengan nama ini).
BRANCH_ORDER = [
    "KLENDER", "CEGER", "BINTARA", "RADJIMAN", "JATIMULYA", "DRAMAGA",
    "CONDET", "JATIBENING", "SAWANGAN", "WARBONG", "CINERE", "CIBINONG",
    "TELUKJ", "JATIWARINGIN", "CIKAMPEK", "CILANGKAP", "PEJATEN", "CIBUBUR",
]
_BRANCH_RANK = {b: i for i, b in enumerate(BRANCH_ORDER)}

BULAN_ID = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
    7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember",
}
BULAN_MAP = {v.lower(): k for k, v in BULAN_ID.items()}
BULAN_ALIAS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "mei": 5, "may": 5, "jun": 6,
    "jul": 7, "agu": 8, "aug": 8, "sep": 9, "okt": 10, "oct": 10, "nov": 11, "des": 12, "dec": 12,
}


def order_branches(names):
    return sorted(set(names), key=lambda b: _BRANCH_RANK.get(str(b).upper(), 999))


# ========================= Format helpers =========================

def format_rupiah(v) -> str:
    try:
        v = float(v)
    except (TypeError, ValueError):
        return "Rp 0"
    sign = "-" if v < 0 else ""
    return f"{sign}Rp {abs(v):,.0f}".replace(",", ".")


def format_number(v) -> str:
    try:
        v = float(v)
    except (TypeError, ValueError):
        return "0"
    return f"{v:,.0f}".replace(",", ".")


def format_decimal(v, digits=1) -> str:
    try:
        v = float(v)
    except (TypeError, ValueError):
        return "0"
    s = f"{v:,.{digits}f}"
    a, _, b = s.partition(".")
    a = a.replace(",", ".")
    return f"{a},{b}" if b else a


def format_percent(v, digits=1) -> str:
    try:
        v = float(v) * 100
    except (TypeError, ValueError):
        return "0%"
    return f"{format_decimal(v, digits)}%"


def sanitize_filename(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", str(name)).strip("_")


def branch_from_filename(fname: str):
    """Deteksi nama cabang dari nama file. Cocok persis dulu; jika tidak ketemu,
    coba cocokkan token setelah 'MFLASH' dengan pencocokan awalan (prefix) dua arah
    supaya tahan terhadap nama cabang yang terpotong pada file export (mis. file
    bernama '...mflashklende_...' tetap terdeteksi sebagai KLENDER)."""
    up = str(fname).upper().replace(" ", "")
    for b in BRANCH_ORDER:
        if b in up:
            return b
    m = re.search(r"MFLASH([A-Z]+)", up)
    if m:
        token = m.group(1)
        for b in BRANCH_ORDER:
            if b.startswith(token) or token.startswith(b):
                return b
    return None


def branch_from_sheetname(sheet: str):
    if not sheet:
        return None
    up = str(sheet).upper().replace(" ", "")
    for b in BRANCH_ORDER:
        if b in up:
            return b
    return None


_SERVICE_KEYWORDS = ["SERVICE", "JASA", "SPAREPART"]


def classify_kategori(v) -> str:
    """Klasifikasi Service vs Gadget & Aksesoris berdasar KATEGORI PENJUALAN."""
    if not v:
        return "Gadget & Aksesoris"
    up = str(v).strip().upper()
    for kw in _SERVICE_KEYWORDS:
        if kw in up:
            return "Service"
    return "Gadget & Aksesoris"


def _extract_filename_timestamp(fname: str):
    """Ekstrak timestamp YYMMDDHHMMSS di akhir nama file (format export MFlash)."""
    m = re.search(r"(\d{12})(?:\.\w+)?$", str(fname))
    if not m:
        m = re.search(r"_(\d{6,14})\.\w+$", str(fname))
        if m:
            return m.group(1)
        return None
    return m.group(1)

# ========================= 6 Pilar MFlash =========================

# 6 Pilar MFlash yang BENAR (dikonfirmasi user) adalah kategori JENIS TRANSAKSI,
# bukan jenis barang: Service, Penjualan Ritel, Sewa, Maintenance, Pengadaan,
# Internet Provider. "Lainnya" adalah bucket tambahan di luar 6 pilar resmi
# untuk transaksi yang tidak bisa diklasifikasikan ke salah satu dari 6 itu.
PILAR_ORDER = ["Service", "Penjualan Ritel", "Sewa", "Maintenance", "Pengadaan", "Internet Provider", "Lainnya"]
PILAR_ICONS = {
    "Service": "🔧", "Penjualan Ritel": "🛒", "Sewa": "🏠",
    "Maintenance": "🛠️", "Pengadaan": "📦", "Internet Provider": "🌐", "Lainnya": "📁",
}
PILAR_COLORS = {
    "Service": "#dc2626", "Penjualan Ritel": "#2563eb", "Sewa": "#7c3aed",
    "Maintenance": "#d97706", "Pengadaan": "#059669", "Internet Provider": "#0891b2", "Lainnya": "#6b7280",
}
_PILAR_SHOW_QTY = {"Penjualan Ritel", "Pengadaan", "Lainnya"}


def _pilar_label(p: str) -> str:
    return p


def _find_pilar_column_index(col_idx: dict):
    """Cari kolom KATEGORI BARANG - dipakai sebagai PROKSI cadangan (bukan
    sumber utama lagi) untuk transaksi yang kolom KATEGORI PILAR-nya kosong/
    tidak spesifik, supaya klasifikasi 6 Pilar tetap lengkap untuk data lama
    (sebelum kolom KATEGORI PILAR ada, mis. sebelum Agustus 2026)."""
    for header, idx in col_idx.items():
        if header == "KATEGORI BARANG":
            return idx
    for header, idx in col_idx.items():
        if "KATEGORI BARANG" in header:
            return idx
    return None


def _find_pilar_excel_column_index(col_idx: dict):
    """Cari kolom 'KATEGORI PILAR ...' asli dari Excel - ini SUMBER UTAMA untuk
    6 Pilar resmi MFlash (Service/Penjualan Ritel/Sewa/Maintenance/Pengadaan/
    Internet Provider). Kolom ini baru mulai diisi sistem MFlash sejak awal
    Agustus 2026, jadi transaksi sebelum itu wajar kosong."""
    for header, idx in col_idx.items():
        if header == "KATEGORI PILAR" or header == "PILAR":
            return idx
    for header, idx in col_idx.items():
        if "PILAR" in header and "BARANG" not in header:
            return idx
    return None


def classify_pilar_official(v) -> str:
    """Klasifikasi 6 Pilar RESMI MFlash langsung dari nilai kolom KATEGORI
    PILAR di Excel (Service, Penjualan Ritel, Sewa, Maintenance, Pengadaan,
    Internet Provider). Kembalikan 'Lainnya' kalau kosong atau tidak cocok
    dengan salah satu dari 6 pilar resmi itu."""
    if not v:
        return "Lainnya"
    up = str(v).strip().upper()
    if "SEWA" in up:
        return "Sewa"
    if "MAINTENANCE" in up or "MAINTAIN" in up:
        return "Maintenance"
    if "INTERNET" in up or "PROVIDER" in up:
        return "Internet Provider"
    if "PENGADAAN" in up:
        return "Pengadaan"
    if "RITEL" in up or "RETAIL" in up:
        return "Penjualan Ritel"
    if "SERVICE" in up:
        return "Service"
    return "Lainnya"


def classify_pilar_barang_proxy(v) -> str:
    """Proksi/perkiraan 6 Pilar dari kolom KATEGORI BARANG, dipakai HANYA
    sebagai fallback ketika KATEGORI PILAR kosong (mis. transaksi sebelum
    Agustus 2026). KATEGORI BARANG cuma bisa membedakan 'ini barang jasa/
    sparepart perbaikan' vs 'ini barang dijual' - jadi hanya bisa menebak
    Service vs Penjualan Ritel; TIDAK bisa mendeteksi Sewa/Maintenance/
    Pengadaan/Internet Provider (bukan tentang jenis barang)."""
    if not v:
        return "Lainnya"
    up = str(v).strip().upper()
    if "JASA" in up or "SPAREPART" in up or "SERVICE" in up:
        return "Service"
    return "Penjualan Ritel"


def classify_pilar_hybrid(pilar_excel_raw, kategori_barang_raw) -> str:
    """Klasifikasi 6 Pilar GABUNGAN (dipakai sebagai default/utama): utamakan
    nilai kolom KATEGORI PILAR asli Excel (sumber resmi). Kalau kosong/tidak
    cocok (mis. transaksi sebelum Agustus 2026 saat kolom itu belum ada),
    fallback ke proksi dari KATEGORI BARANG supaya hasilnya tetap lengkap."""
    excel_cls = classify_pilar_official(pilar_excel_raw)
    if excel_cls != "Lainnya":
        return excel_cls
    return classify_pilar_barang_proxy(kategori_barang_raw)


def parse_bulan(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        iv = int(v)
        return iv if 1 <= iv <= 12 else None
    s = str(v).strip().lower()
    if s in BULAN_MAP:
        return BULAN_MAP[s]
    if s in BULAN_ALIAS:
        return BULAN_ALIAS[s]
    for alias, num in BULAN_ALIAS.items():
        if s.startswith(alias):
            return num
    return None


def _nan_to_none(v):
    """Konversi NaN/NaT pandas jadi None. Penting karena bool(float('nan')) True
    di Python, sehingga NaN bisa lolos pengecekan 'if v' dan merusak fallback
    logic maupun konversi angka."""
    if v is None:
        return None
    try:
        if isinstance(v, float) and pd.isna(v):
            return None
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    return v


def to_date(v):
    v = _nan_to_none(v)
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)):
        try:
            return (datetime(1899, 12, 30) + timedelta(days=float(v))).date()
        except (OverflowError, ValueError):
            return None
    s = str(v).strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d", "%d %B %Y", "%d %b %Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    try:
        return pd.to_datetime(s, errors="coerce").date()
    except Exception:
        return None


def _to_float_or_none(v):
    v = _nan_to_none(v)
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        try:
            return float(str(v).strip())
        except ValueError:
            return None


def _build_col_idx(header_row) -> dict:
    col_idx = {}
    for i, h in enumerate(header_row):
        h = _nan_to_none(h)
        if h is None:
            continue
        key = str(h).strip().upper()
        if key and key not in col_idx:
            col_idx[key] = i
    return col_idx


def _find_penjual_column_index(col_idx: dict):
    for header, idx in col_idx.items():
        if "DEFAULT PENJUAL" in header:
            return idx
    for header, idx in col_idx.items():
        if "PENJUAL" in header:
            return idx
    return None


# ========================= Marketing Corporate classification =========================

_MC_KEYWORDS = ["CORPORATE", "CORP", "MARKETING CORPORATE", "MC"]


def classify_mc_or_retail(kategori_pelanggan) -> str:
    if not kategori_pelanggan:
        return "Retail"
    up = str(kategori_pelanggan).strip().upper()
    for kw in _MC_KEYWORDS:
        if kw in up:
            return "Marketing Corporate"
    return "Retail"

# ========================= Loader Data Omset (Main) =========================

def _looks_like_ads_export(path: str) -> bool:
    try:
        xls = pd.ExcelFile(path)
        for name in xls.sheet_names:
            up = name.upper()
            if "CAMPAIGN" in up or "AD SET" in up or "ADS" in up:
                return True
        raw = pd.read_excel(path, sheet_name=xls.sheet_names[0], header=None, nrows=1)
        header_row = [str(_nan_to_none(v) or "") for v in list(raw.iloc[0])]
        joined = " ".join(header_row).upper()
        return "CAMPAIGN NAME" in joined or "AMOUNT SPENT" in joined
    except Exception:
        return False


def _detect_main_file_kind(path: str) -> str:
    try:
        xls = pd.ExcelFile(path)
        names_upper = [n.upper() for n in xls.sheet_names]
        for n in names_upper:
            if "RINCIAN FAKTUR" in n or "FAKTUR PENJUALAN" in n:
                return "faktur"
        if MAIN_SHEET_NAME.upper() in names_upper or "SCOREBOARD" in names_upper:
            return "master"
    except Exception:
        pass
    return "faktur"


def _load_faktur_sheet(path: str, cabang_hint=None) -> pd.DataFrame:
    """Loader untuk format baru: 'Rincian Faktur Penjualan' per cabang.
    Pakai pandas.read_excel (bukan openpyxl read_only) karena sejumlah file
    export MFlash punya metadata dimensi sheet yang tidak akurat sehingga
    openpyxl read_only gagal mendeteksi baris data (sama seperti bug pada
    loader Walk-in)."""
    try:
        sheet_name = None
        xls = pd.ExcelFile(path)
        for name in xls.sheet_names:
            if "RINCIAN FAKTUR" in name.upper() or "FAKTUR PENJUALAN" in name.upper():
                sheet_name = name
                break
        if sheet_name is None:
            sheet_name = xls.sheet_names[0]
        raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
    except Exception:
        return pd.DataFrame()
    if raw is None or raw.empty:
        return pd.DataFrame()

    header_row = list(raw.iloc[0])
    col_idx = _build_col_idx(header_row)
    rows_iter = (tuple(r) for r in raw.iloc[1:].itertuples(index=False, name=None))

    def gi(*names):
        for n in names:
            if n in col_idx:
                return col_idx[n]
        return None

    idx_cabang = gi("CABANG")
    idx_tgl = gi("TGL FAKTUR", "TANGGAL")
    idx_kategori = gi("KATEGORI PENJUALAN")
    idx_kategori_pelanggan = gi("KATEGORI PELANGGAN")
    idx_total = gi("TOTAL HARGA")
    idx_qty = gi("QTY")
    idx_gp = gi("GROSS PROFIT")
    idx_pilar = _find_pilar_column_index(col_idx)
    idx_pilar_excel = _find_pilar_excel_column_index(col_idx)
    idx_penjual = _find_penjual_column_index(col_idx)

    cabang_fallback = cabang_hint or branch_from_filename(os.path.basename(path)) or branch_from_sheetname(sheet_name)

    records = []
    for row in rows_iter:
        if row is None:
            continue
        total = _to_float_or_none(row[idx_total]) if idx_total is not None and idx_total < len(row) else None
        if total is None:
            continue
        cabang = _nan_to_none(row[idx_cabang]) if idx_cabang is not None and idx_cabang < len(row) else None
        cabang = str(cabang).strip().upper() if cabang else cabang_fallback
        tgl = to_date(row[idx_tgl]) if idx_tgl is not None and idx_tgl < len(row) else None
        kategori_raw = _nan_to_none(row[idx_kategori]) if idx_kategori is not None and idx_kategori < len(row) else None
        kategori_pelanggan_raw = _nan_to_none(row[idx_kategori_pelanggan]) if idx_kategori_pelanggan is not None and idx_kategori_pelanggan < len(row) else None
        pilar_raw = _nan_to_none(row[idx_pilar]) if idx_pilar is not None and idx_pilar < len(row) else None
        pilar_excel_raw = _nan_to_none(row[idx_pilar_excel]) if idx_pilar_excel is not None and idx_pilar_excel < len(row) else None
        penjual_raw = _nan_to_none(row[idx_penjual]) if idx_penjual is not None and idx_penjual < len(row) else None
        qty = _to_float_or_none(row[idx_qty]) if idx_qty is not None and idx_qty < len(row) else 0.0
        gp = _to_float_or_none(row[idx_gp]) if idx_gp is not None and idx_gp < len(row) else 0.0
        records.append({
            "Cabang": cabang,
            "Tanggal": tgl,
            "Kategori": classify_kategori(kategori_raw),
            "Omset": total,
            "Qty": qty or 0.0,
            "GrossProfit": gp or 0.0,
            "Pilar": classify_pilar_hybrid(pilar_excel_raw, pilar_raw),
            "PilarExcel": classify_pilar_official(pilar_excel_raw),
            "PilarSource": "Excel" if classify_pilar_official(pilar_excel_raw) != "Lainnya" else "Barang",
            "NamaPenjual": str(penjual_raw).strip() if penjual_raw else "TIDAK DIKETAHUI",
            "PenjualKelompok": classify_mc_or_retail(kategori_pelanggan_raw),
        })
    if not records:
        return pd.DataFrame()
    return pd.DataFrame(records)


def _load_master_sheet(path: str) -> pd.DataFrame:
    """Loader untuk format lama: file master dengan sheet 'Faktur Penjualan'."""
    try:
        xls = pd.ExcelFile(path)
        sheet_name = None
        for name in xls.sheet_names:
            if name.upper() == MAIN_SHEET_NAME.upper():
                sheet_name = name
                break
        if sheet_name is None:
            return pd.DataFrame()
        raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
    except Exception:
        return pd.DataFrame()
    if raw is None or raw.empty:
        return pd.DataFrame()

    header_row = list(raw.iloc[0])
    col_idx = _build_col_idx(header_row)
    rows_iter = (tuple(r) for r in raw.iloc[1:].itertuples(index=False, name=None))

    def gi(*names):
        for n in names:
            if n in col_idx:
                return col_idx[n]
        return None

    idx_cabang = gi("CABANG")
    idx_tgl = gi("TGL FAKTUR", "TANGGAL")
    idx_kategori = gi("KATEGORI PENJUALAN")
    idx_kategori_pelanggan = gi("KATEGORI PELANGGAN")
    idx_total = gi("TOTAL HARGA")
    idx_qty = gi("QTY")
    idx_gp = gi("GROSS PROFIT")
    idx_pilar = _find_pilar_column_index(col_idx)
    idx_pilar_excel = _find_pilar_excel_column_index(col_idx)
    idx_penjual = _find_penjual_column_index(col_idx)

    records = []
    for row in rows_iter:
        if row is None:
            continue
        total = _to_float_or_none(row[idx_total]) if idx_total is not None and idx_total < len(row) else None
        if total is None:
            continue
        cabang = _nan_to_none(row[idx_cabang]) if idx_cabang is not None and idx_cabang < len(row) else None
        cabang = str(cabang).strip().upper() if cabang else None
        tgl = to_date(row[idx_tgl]) if idx_tgl is not None and idx_tgl < len(row) else None
        kategori_raw = _nan_to_none(row[idx_kategori]) if idx_kategori is not None and idx_kategori < len(row) else None
        kategori_pelanggan_raw = _nan_to_none(row[idx_kategori_pelanggan]) if idx_kategori_pelanggan is not None and idx_kategori_pelanggan < len(row) else None
        pilar_raw = _nan_to_none(row[idx_pilar]) if idx_pilar is not None and idx_pilar < len(row) else None
        pilar_excel_raw = _nan_to_none(row[idx_pilar_excel]) if idx_pilar_excel is not None and idx_pilar_excel < len(row) else None
        penjual_raw = _nan_to_none(row[idx_penjual]) if idx_penjual is not None and idx_penjual < len(row) else None
        qty = _to_float_or_none(row[idx_qty]) if idx_qty is not None and idx_qty < len(row) else 0.0
        gp = _to_float_or_none(row[idx_gp]) if idx_gp is not None and idx_gp < len(row) else 0.0
        records.append({
            "Cabang": cabang,
            "Tanggal": tgl,
            "Kategori": classify_kategori(kategori_raw),
            "Omset": total,
            "Qty": qty or 0.0,
            "GrossProfit": gp or 0.0,
            "Pilar": classify_pilar_hybrid(pilar_excel_raw, pilar_raw),
            "PilarExcel": classify_pilar_official(pilar_excel_raw),
            "PilarSource": "Excel" if classify_pilar_official(pilar_excel_raw) != "Lainnya" else "Barang",
            "NamaPenjual": str(penjual_raw).strip() if penjual_raw else "TIDAK DIKETAHUI",
            "PenjualKelompok": classify_mc_or_retail(kategori_pelanggan_raw),
        })
    if not records:
        return pd.DataFrame()
    return pd.DataFrame(records)


def load_main_data(path: str) -> pd.DataFrame:
    kind = _detect_main_file_kind(path)
    if kind == "faktur":
        return _load_faktur_sheet(path)
    if kind == "master":
        return _load_master_sheet(path)
    try:
        return _load_faktur_sheet(path)
    except Exception:
        return pd.DataFrame()


@st.cache_data(show_spinner=False)
def _load_main_data_cached(path: str, mtime: float, size: int) -> pd.DataFrame:
    """Wrapper cache: file Excel yang sama (path+mtime+size tidak berubah) tidak
    akan dibaca & diparse ulang setiap kali Streamlit rerun script (setiap ada
    interaksi seperti ganti tanggal/filter/klik tombol)."""
    return load_main_data(path)


def _dedupe_main_files():
    """Hapus file Omset duplikat/basi per cabang, sisakan yang timestamp-nya terbaru."""
    if not os.path.isdir(MAIN_DATA_DIR):
        return
    files = [f for f in os.listdir(MAIN_DATA_DIR) if f.lower().endswith((".xlsx", ".xls"))]
    groups = {}
    for f in files:
        branch = branch_from_filename(f) or "UNKNOWN"
        ts = _extract_filename_timestamp(f) or ""
        groups.setdefault(branch, []).append((ts, f))
    for branch, items in groups.items():
        if branch == "UNKNOWN" or len(items) <= 1:
            continue
        items.sort(key=lambda x: x[0])
        stale = items[:-1]
        for ts, fname in stale:
            fpath = os.path.join(MAIN_DATA_DIR, fname)
            try:
                os.remove(fpath)
                if _GH_ENABLED:
                    github_delete_file(f"data/main/{fname}", "auto-dedupe stale branch file")
            except Exception:
                pass


def load_all_main_data() -> pd.DataFrame:
    if not os.path.isdir(MAIN_DATA_DIR):
        return pd.DataFrame()
    # PENTING: jalankan dedupe di SETIAP load (bukan cuma saat upload baru).
    # File export MFlash bersifat KUMULATIF PENUH dari awal kuartal setiap
    # kali di-export ulang (bukan cuma data baru/incremental) - jadi kalau
    # ada 2 file untuk cabang yang sama (mis. sisa dari restore GitHub, atau
    # dedupe sebelumnya sempat gagal), transaksi yang tanggalnya tumpang
    # tindih akan TERHITUNG DOBEL dan bikin Omset S/D Hari Ini jadi lebih
    # besar dari yang sebenarnya. Jalankan di sini supaya selalu bersih
    # sebelum data digabung & di-cache, apapun penyebab file duplikatnya.
    _dedupe_main_files()
    cached = _load_cached_combined(MAIN_DATA_DIR, "main_combined", required_cols=["Pilar", "PilarExcel", "PilarSource"], schema_version=MAIN_DATA_SCHEMA_VERSION)
    if cached is not None:
        return cached
    frames = []
    for fname in sorted(os.listdir(MAIN_DATA_DIR)):
        if not fname.lower().endswith((".xlsx", ".xls")):
            continue
        fpath = os.path.join(MAIN_DATA_DIR, fname)
        try:
            stat_ = os.stat(fpath)
            df = _load_main_data_cached(fpath, stat_.st_mtime, stat_.st_size)
            if not df.empty:
                frames.append(df)
        except Exception:
            continue
    if not frames:
        return pd.DataFrame()
    combined = pd.concat(frames, ignore_index=True)
    combined["Tahun"] = combined["Tanggal"].apply(lambda d: d.year if d else None)
    combined["Bulan"] = combined["Tanggal"].apply(lambda d: d.month if d else None)
    combined = combined.dropna(subset=["Cabang", "Tanggal"])
    _save_cached_combined(MAIN_DATA_DIR, "main_combined", combined, schema_version=MAIN_DATA_SCHEMA_VERSION)
    return combined

# ========================= Loader Data Iklan (Meta Ads) =========================

_ADS_REQUIRED_COLS = ["Campaign name", "Amount spent", "Reach", "Impressions"]


def _dedupe_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Buang kolom duplikat (nama sama setelah rename) supaya pd.concat tidak
    error InvalidIndexError."""
    if df is None or df.empty:
        return df
    return df.loc[:, ~df.columns.duplicated()]


_ADS_RENAME_MAP = {
    "campaign name": "CampaignName",
    "amount spent (idr)": "AmountSpent",
    "amount spent": "AmountSpent",
    "reach": "Reach",
    "impressions": "Impressions",
    "link clicks": "Clicks",
    "clicks (all)": "Clicks",
    "cpm (cost per 1,000 impressions) (idr)": "CPM",
    "cpm (cost per 1000 impressions)": "CPM",
    "cpc (cost per link click) (idr)": "CPC",
    "cpc (all) (idr)": "CPC",
    "ctr (link click-through rate)": "CTR",
    "ctr (all)": "CTR",
    "results": "Results",
    "messaging conversations started": "Results",
}


def load_ads_data(path: str, cabang_hint=None) -> pd.DataFrame:
    try:
        if path.lower().endswith(".csv"):
            df = pd.read_csv(path)
        else:
            df = pd.read_excel(path)
    except Exception:
        return pd.DataFrame()
    if df is None or df.empty:
        return pd.DataFrame()
    df.columns = [str(c).strip() for c in df.columns]
    rename = {}
    for c in df.columns:
        key = c.strip().lower()
        if key in _ADS_RENAME_MAP:
            rename[c] = _ADS_RENAME_MAP[key]
    df = df.rename(columns=rename)
    df = _dedupe_columns(df)
    cabang = cabang_hint or branch_from_filename(os.path.basename(path))
    df["Cabang"] = cabang or "TIDAK DIKETAHUI"
    for c in ["AmountSpent", "Reach", "Impressions", "Clicks", "Results"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


@st.cache_data(show_spinner=False)
def _load_ads_data_cached(path: str, mtime: float, size: int) -> pd.DataFrame:
    """Wrapper cache untuk loader Iklan (path+mtime+size sebagai key) supaya
    file yang belum berubah tidak diparse ulang di setiap rerun."""
    return load_ads_data(path)


def load_all_ads_data() -> pd.DataFrame:
    if not os.path.isdir(ADS_DATA_DIR):
        return pd.DataFrame()
    cached = _load_cached_combined(ADS_DATA_DIR, "ads_combined")
    if cached is not None:
        return cached
    frames = []
    for fname in sorted(os.listdir(ADS_DATA_DIR)):
        if not fname.lower().endswith((".xlsx", ".xls", ".csv")):
            continue
        fpath = os.path.join(ADS_DATA_DIR, fname)
        try:
            stat_ = os.stat(fpath)
            df = _load_ads_data_cached(fpath, stat_.st_mtime, stat_.st_size)
            if df is not None and not df.empty:
                frames.append(_dedupe_columns(df))
        except Exception:
            continue
    if not frames:
        return pd.DataFrame()
    try:
        result = pd.concat(frames, ignore_index=True, sort=False)
    except Exception:
        keep_cols = ["Cabang", "CampaignName", "AmountSpent", "Reach", "Impressions",
                     "Clicks", "CPM", "CPC", "CTR", "Results"]
        cleaned = []
        for f in frames:
            f = _dedupe_columns(f)
            cols_present = [c for c in keep_cols if c in f.columns]
            cleaned.append(f[cols_present])
        result = pd.concat(cleaned, ignore_index=True, sort=False)
    _save_cached_combined(ADS_DATA_DIR, "ads_combined", result)
    return result


def aggregate_ads_by_branch(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty or "Cabang" not in df.columns:
        return pd.DataFrame()
    agg_map = {}
    for c in ["AmountSpent", "Reach", "Impressions", "Clicks", "Results"]:
        if c in df.columns:
            agg_map[c] = "sum"
    if not agg_map:
        return pd.DataFrame()
    g = df.groupby("Cabang").agg(agg_map).reset_index()
    if "AmountSpent" in g.columns and "Results" in g.columns:
        g["CostPerResult"] = g.apply(lambda r: (r["AmountSpent"] / r["Results"]) if r["Results"] else None, axis=1)
    return g


def generate_ads_insights(df_branch: pd.DataFrame):
    insights = []
    if df_branch is None or df_branch.empty:
        return insights
    if "CostPerResult" in df_branch.columns:
        valid = df_branch.dropna(subset=["CostPerResult"])
        if not valid.empty:
            worst = valid.sort_values("CostPerResult", ascending=False).iloc[0]
            insights.append({
                "title": f"Cost per Result tertinggi: {worst['Cabang']}",
                "detail": f"Rp {format_number(worst['CostPerResult'])} per hasil - evaluasi kreatif/targeting iklan cabang ini.",
                "level": "warning",
            })
    return insights

# ========================= Sales Insight Engine =========================

def render_structured_insight_card(ins: dict):
    level = ins.get("level", "info")
    colors = {"info": "#2563eb", "warning": "#d97706", "danger": "#dc2626", "success": "#16a34a"}
    icons = {"info": "ℹ️", "warning": "⚠️", "danger": "🔴", "success": "✅"}
    color = colors.get(level, "#2563eb")
    icon = icons.get(level, "ℹ️")
    st.markdown(
        (
            f'<div style="background:white;border-left:4px solid {color};border-radius:8px;'
            f'padding:12px 16px;margin-bottom:10px;box-shadow:0 1px 2px rgba(0,0,0,0.06);">'
            f'<div style="font-weight:700;color:{color};">{icon} {ins.get("title","")}</div>'
            f'<div style="color:#374151;font-size:0.92em;margin-top:4px;">{ins.get("detail","")}</div>'
            f'</div>'
        ),
        unsafe_allow_html=True,
    )


def generate_sales_insights(df: pd.DataFrame, kategori_label: str):
    insights = []
    if df is None or df.empty:
        return insights
    by_branch = df.groupby("Cabang")["Omset"].sum().reset_index().sort_values("Omset")
    if not by_branch.empty:
        lowest = by_branch.iloc[0]
        insights.append({
            "title": f"{kategori_label} terendah: {lowest['Cabang']}",
            "detail": f"Omset {kategori_label} cabang {lowest['Cabang']} sebesar {format_rupiah(lowest['Omset'])} - "
                      f"perlu evaluasi. Rencana aksi: (online) tingkatkan promosi digital lokal & respons cepat "
                      f"chat masuk; (offline) aktifkan sales lapangan & program referral pelanggan.",
            "level": "warning",
        })
    return insights


def generate_all_sales_insights(df: pd.DataFrame):
    insights = []
    if df is None or df.empty:
        return insights
    for kategori in df["Kategori"].dropna().unique():
        sub = df[df["Kategori"] == kategori]
        insights.extend(generate_sales_insights(sub, kategori))
    return insights


def render_kpi_card(title: str, value: str, color: str, icon: str = "") -> str:
    return (
        f'<div style="background:white;border:2px solid {color};border-radius:14px;'
        f'padding:20px;text-align:center;box-shadow:0 1px 3px rgba(0,0,0,0.08);">'
        f'<div style="font-size:1.6em;">{icon}</div>'
        f'<div style="color:{color};font-weight:700;font-size:0.95em;margin-top:6px;">{title}</div>'
        f'<div style="font-size:1.5em;font-weight:800;color:#111827;margin-top:4px;">{value}</div>'
        f'</div>'
    )

# ========================= Loader Data Walk-in =========================

def load_walkin_data(path: str, cabang_hint=None) -> pd.DataFrame:
    """Loader Walk-in pakai pandas.read_excel (bukan openpyxl read_only) karena
    file export MFlash punya metadata dimensi sheet yang kadang tidak akurat,
    yang bikin openpyxl read_only gagal mendeteksi baris data (return kosong)."""
    try:
        xls = pd.ExcelFile(path)
        sheet_name = xls.sheet_names[0]
        raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
    except Exception:
        return pd.DataFrame()
    if raw is None or raw.empty:
        return pd.DataFrame()

    header_row = list(raw.iloc[0])
    col_idx = _build_col_idx(header_row)
    rows_iter = (tuple(r) for r in raw.iloc[1:].itertuples(index=False, name=None))

    def gi(*names):
        for n in names:
            if n in col_idx:
                return col_idx[n]
        return None

    idx_cabang = gi("CABANG")
    idx_tgl = gi("TANGGAL", "TGL PENGIRIMAN", "TGL")
    idx_nomor = gi("NOMOR PENGIRIMAN", "NO PENGIRIMAN", "NOMOR")

    cabang_fallback = cabang_hint or branch_from_filename(os.path.basename(path)) or branch_from_sheetname(sheet_name)

    records = []
    for row in rows_iter:
        if row is None:
            continue
        nomor = _nan_to_none(row[idx_nomor]) if idx_nomor is not None and idx_nomor < len(row) else None
        if nomor is None:
            continue
        cabang = _nan_to_none(row[idx_cabang]) if idx_cabang is not None and idx_cabang < len(row) else None
        cabang = str(cabang).strip().upper() if cabang else cabang_fallback
        tgl = to_date(row[idx_tgl]) if idx_tgl is not None and idx_tgl < len(row) else None
        records.append({
            "Cabang": cabang,
            "Tanggal": tgl,
            "NomorPengiriman": str(nomor).strip(),
        })
    if not records:
        return pd.DataFrame()
    df = pd.DataFrame(records)
    df["Tahun"] = df["Tanggal"].apply(lambda d: d.year if d else None)
    df["Bulan"] = df["Tanggal"].apply(lambda d: d.month if d else None)
    return df


@st.cache_data(show_spinner=False)
def _load_walkin_data_cached(path: str, mtime: float, size: int) -> pd.DataFrame:
    """Wrapper cache untuk loader Walk-in (path+mtime+size sebagai key) supaya
    file yang belum berubah tidak diparse ulang di setiap rerun."""
    return load_walkin_data(path)


def load_all_walkin_data() -> pd.DataFrame:
    if not os.path.isdir(WALKIN_DATA_DIR):
        return pd.DataFrame()
    cached = _load_cached_combined(WALKIN_DATA_DIR, "walkin_combined")
    if cached is not None:
        return cached
    frames = []
    for fname in sorted(os.listdir(WALKIN_DATA_DIR)):
        if not fname.lower().endswith((".xlsx", ".xls")):
            continue
        fpath = os.path.join(WALKIN_DATA_DIR, fname)
        try:
            stat_ = os.stat(fpath)
            df = _load_walkin_data_cached(fpath, stat_.st_mtime, stat_.st_size)
            if not df.empty:
                frames.append(df)
        except Exception:
            continue
    if not frames:
        return pd.DataFrame()
    combined = pd.concat(frames, ignore_index=True)
    combined = combined.dropna(subset=["Cabang", "Tanggal"])
    combined = combined.drop_duplicates(subset=["Cabang", "NomorPengiriman"])
    _save_cached_combined(WALKIN_DATA_DIR, "walkin_combined", combined)
    return combined


def aggregate_walkin_monthly(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Cabang", "Tahun", "Bulan", "TotalWalkin", "RataRataPerHari"])
    rows = []
    for (cabang, tahun, bulan), g in df.groupby(["Cabang", "Tahun", "Bulan"]):
        total = g["NomorPengiriman"].nunique()
        hari_dalam_bulan = calendar.monthrange(int(tahun), int(bulan))[1]
        rata2 = total / hari_dalam_bulan if hari_dalam_bulan else 0.0
        rows.append({"Cabang": cabang, "Tahun": tahun, "Bulan": bulan, "TotalWalkin": int(total), "RataRataPerHari": rata2})
    return pd.DataFrame(rows)


def _quarter_bounds_for(d: date):
    q_start_month = ((d.month - 1) // 3) * 3 + 1
    start = date(d.year, q_start_month, 1)
    if q_start_month == 10:
        end = date(d.year, 12, 31)
    else:
        end = date(d.year, q_start_month + 3, 1) - timedelta(days=1)
    total_hari = (end - start).days + 1
    hari_berjalan = (d - start).days + 1
    sisa_hari = max(total_hari - hari_berjalan, 0)
    return start, end, total_hari, hari_berjalan, sisa_hari


def aggregate_walkin_current_period(df: pd.DataFrame, tanggal_acuan: date) -> pd.DataFrame:
    """Total Walk-in KUMULATIF dari awal kuartal (1 Juli/Okt/Jan/Apr) sampai
    tanggal_acuan (inklusif) - konsisten dengan S/D HARI INI di Scoreboard."""
    if df.empty:
        return pd.DataFrame(columns=["Cabang", "TotalWalkin"])
    q_start, q_end, _, _, _ = _quarter_bounds_for(tanggal_acuan)
    mask = (df["Tanggal"] >= q_start) & (df["Tanggal"] <= tanggal_acuan)
    sub = df[mask]
    if sub.empty:
        return pd.DataFrame(columns=["Cabang", "TotalWalkin"])
    g = sub.groupby("Cabang")["NomorPengiriman"].nunique().reset_index()
    g.columns = ["Cabang", "TotalWalkin"]
    return g


def _walkin_ordered(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    df["_rank"] = df["Cabang"].apply(lambda b: _BRANCH_RANK.get(str(b).upper(), 999))
    df = df.sort_values("_rank").drop(columns=["_rank"])
    return df


def _walkin_overall_avg(df_summary: pd.DataFrame) -> float:
    if df_summary.empty or "TotalWalkin" not in df_summary.columns:
        return 0.0
    return float(df_summary["TotalWalkin"].mean())

# ========================= Walk-in render/export/insight =========================

def render_walkin_table_html(df_summary: pd.DataFrame, overall_avg: float) -> str:
    if df_summary.empty:
        return "<p style='color:#6b7280;'>Belum ada data Walk-in.</p>"
    rows_html = ""
    for _, r in df_summary.iterrows():
        total = r["TotalWalkin"]
        color = "#16a34a" if total >= overall_avg else "#dc2626"
        rows_html += f"""<tr>
        <td style="padding:8px 12px;border:1px solid #e5e7eb;">{r['Cabang']}</td>
        <td style="padding:8px 12px;border:1px solid #e5e7eb;text-align:right;color:{color};font-weight:700;">{format_number(total)}</td>
        </tr>"""
    return f"""<table style="width:100%;border-collapse:collapse;">
    <thead><tr style="background:#f3f4f6;">
    <th style="padding:8px 12px;border:1px solid #e5e7eb;text-align:left;">Cabang</th>
    <th style="padding:8px 12px;border:1px solid #e5e7eb;text-align:right;">Total Walk-in</th>
    </tr></thead><tbody>{rows_html}</tbody></table>"""


def generate_walkin_table_image(df_summary: pd.DataFrame, title: str = "Walk-in per Cabang") -> bytes:
    fig, ax = plt.subplots(figsize=(6, max(2, 0.4 * len(df_summary) + 1)))
    ax.axis("off")
    if df_summary.empty:
        ax.text(0.5, 0.5, "Tidak ada data", ha="center", va="center")
    else:
        table_data = [[r["Cabang"], format_number(r["TotalWalkin"])] for _, r in df_summary.iterrows()]
        tbl = ax.table(cellText=table_data, colLabels=["Cabang", "Total Walk-in"], loc="center", cellLoc="center")
        tbl.auto_set_font_size(False)
        tbl.set_fontsize(9)
        tbl.scale(1, 1.4)
    ax.set_title(title, fontweight="bold")
    buf = io.BytesIO()
    fig.savefig(buf, format="jpg", dpi=150, bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


def generate_walkin_table_pdf(df_summary: pd.DataFrame, title: str = "Walk-in per Cabang") -> bytes:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as pdf_canvas
    buf = io.BytesIO()
    c = pdf_canvas.Canvas(buf, pagesize=A4)
    width, height = A4
    c.setFont("Helvetica-Bold", 14)
    c.drawString(40, height - 50, title)
    c.setFont("Helvetica", 10)
    y = height - 90
    for _, r in df_summary.iterrows():
        c.drawString(40, y, str(r["Cabang"]))
        c.drawRightString(300, y, format_number(r["TotalWalkin"]))
        y -= 18
        if y < 60:
            c.showPage()
            y = height - 50
    c.save()
    return buf.getvalue()


_WALKIN_ACTION_PLAN = (
    "Rencana aksi: (online) optimalkan Google Maps/Ads lokal & konten media sosial cabang; "
    "(offline) pasang spanduk/flyer area sekitar & aktifkan sales canvassing ke perkantoran/perumahan terdekat."
)


def generate_walkin_marketing_insights(df_summary: pd.DataFrame, overall_avg: float):
    insights = []
    if df_summary.empty:
        return insights
    below = df_summary[df_summary["TotalWalkin"] < overall_avg].sort_values("TotalWalkin")
    for _, r in below.head(3).iterrows():
        insights.append({
            "title": f"Walk-in rendah: {r['Cabang']}",
            "detail": f"Total Walk-in {format_number(r['TotalWalkin'])} di bawah rata-rata cabang lain "
                      f"({format_number(overall_avg)}). {_WALKIN_ACTION_PLAN}",
            "level": "warning",
        })
    return insights


def generate_walkin_insights(df_summary: pd.DataFrame):
    overall_avg = _walkin_overall_avg(df_summary)
    return generate_walkin_marketing_insights(df_summary, overall_avg)

# ========================= Corporate/Target loaders =========================

def load_corporate_data(path: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
    except Exception:
        return pd.DataFrame()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def make_corporate_template() -> bytes:
    df = pd.DataFrame(columns=["Cabang", "Nama Sales", "Target Bulan Ini", "S/D Hari Ini"])
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Corporate")
    return buf.getvalue()


SCOREBOARD_KATEGORI = ["Omset All", "Service", "Gadget & Aksesoris"]


def load_target_data(path: str):
    """Baca file Target Omset -> dict {kategori: {cabang: target}}.
    Mendukung DUA format:
    1. Format LEBAR (paling umum dipakai user secara natural): satu baris per
       cabang, satu kolom per kategori - mis. kolom 'Cabang', 'Target Service',
       'Target Gadget', 'Target All', (opsional) 'Target Corporate'.
    2. Format PANJANG (sesuai template bawaan dashboard): kolom 'Cabang',
       'Kategori', 'Target' - satu baris per kombinasi cabang x kategori.
    """
    target_map = {k: {} for k in SCOREBOARD_KATEGORI}
    try:
        df = pd.read_excel(path, sheet_name=0)
    except Exception:
        return target_map
    df.columns = [str(c).strip() for c in df.columns]
    col_cabang = next((c for c in df.columns if c.strip().upper() == "CABANG"), None)
    if col_cabang is None:
        return target_map

    # --- Coba format LEBAR dulu: cari semua kolom yang mengandung kata "TARGET"
    # lalu petakan ke kategori scoreboard berdasar kata kunci di nama kolomnya.
    wide_col_map = {}
    for c in df.columns:
        if c == col_cabang:
            continue
        up = c.strip().upper()
        if "TARGET" not in up:
            continue
        if "SERVICE" in up:
            wide_col_map[c] = "Service"
        elif "GADGET" in up or "AKSESORIS" in up or "ACCESSORIES" in up:
            wide_col_map[c] = "Gadget & Aksesoris"
        elif "CORPORATE" in up or "CORP" in up:
            wide_col_map[c] = "Marketing Corporate"
        elif "ALL" in up or up.strip() == "TARGET":
            wide_col_map[c] = "Omset All"

    if wide_col_map:
        for _, row in df.iterrows():
            cabang = _nan_to_none(row.get(col_cabang))
            if not cabang:
                continue
            cabang = str(cabang).strip().upper()
            for col, kategori in wide_col_map.items():
                target = _to_float_or_none(row.get(col))
                if target is None:
                    continue
                if kategori not in target_map:
                    target_map[kategori] = {}
                target_map[kategori][cabang] = target
        return target_map

    # --- Fallback ke format PANJANG (Cabang, Kategori, Target).
    col_kategori = next((c for c in df.columns if c.strip().upper() == "KATEGORI"), None)
    col_target = next((c for c in df.columns if "TARGET" in c.strip().upper()), None)
    if not (col_kategori and col_target):
        return target_map
    for _, row in df.iterrows():
        cabang = _nan_to_none(row.get(col_cabang))
        kategori = _nan_to_none(row.get(col_kategori))
        target = _to_float_or_none(row.get(col_target))
        if not (cabang and kategori and target is not None):
            continue
        cabang = str(cabang).strip().upper()
        kategori = str(kategori).strip()
        if kategori not in target_map:
            target_map[kategori] = {}
        target_map[kategori][cabang] = target
    return target_map


def make_target_template() -> bytes:
    rows = []
    for b in BRANCH_ORDER:
        for k in SCOREBOARD_KATEGORI:
            rows.append({"Cabang": b, "Kategori": k, "Target Kuartal": 0})
    df = pd.DataFrame(rows)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Target")
    return buf.getvalue()


# ========================= Scoreboard core =========================

def _quarter_bounds(d: date):
    """Kembalikan (start, end, total_hari, hari_berjalan, sisa_hari) untuk
    kuartal kalender (Jan-Mar, Apr-Jun, Jul-Sep, Okt-Des) yang memuat tanggal d."""
    q_start_month = ((d.month - 1) // 3) * 3 + 1
    start = date(d.year, q_start_month, 1)
    if q_start_month == 10:
        end = date(d.year, 12, 31)
    else:
        end = date(d.year, q_start_month + 3, 1) - timedelta(days=1)
    total_hari = (end - start).days + 1
    hari_berjalan = (d - start).days + 1
    sisa_hari = max(total_hari - hari_berjalan, 0)
    return start, end, total_hari, hari_berjalan, sisa_hari


# alias dipakai oleh loader Walk-in (definisi sama persis)
_quarter_bounds_for = _quarter_bounds


def target_period_caption(tanggal_acuan: date) -> str:
    """Teks penjelasan bahwa Target Omset berlaku untuk 1 kuartal PENUH
    (mis. Jul-Sep), tidak perlu upload ulang setiap hari - hanya perlu
    upload ulang saat masuk kuartal berikutnya dengan angka target baru."""
    start, end, _, _, _ = _quarter_bounds(tanggal_acuan)
    return (f"Target Omset berlaku untuk 1 kuartal penuh: **{start.strftime('%d %b')} - {end.strftime('%d %b %Y')}**. "
            f"Tidak perlu upload ulang setiap hari — cukup upload sekali di awal kuartal ini, dan upload lagi "
            f"dengan angka baru saat masuk kuartal berikutnya.")


def pencapaian_color(pct):
    """Hijau >=100%, kuning 85-99.9%, merah <85%, abu-abu kalau tidak ada Target (None)."""
    if pct is None:
        return "#9ca3af"
    if pct >= 1.0:
        return "#16a34a"
    if pct >= 0.85:
        return "#d97706"
    return "#dc2626"


def _kategori_filter(df_main: pd.DataFrame, kategori: str) -> pd.DataFrame:
    if kategori == "Omset All":
        return df_main
    return df_main[df_main["Kategori"] == kategori]


def build_scoreboard(df_main: pd.DataFrame, target_map: dict, tanggal_acuan: date,
                      branches, kategori: str) -> pd.DataFrame:
    start, end, total_hari, hari_berjalan, sisa_hari = _quarter_bounds(tanggal_acuan)
    sub = _kategori_filter(df_main, kategori) if not df_main.empty else df_main
    rows = []
    kat_targets = target_map.get(kategori, {}) if target_map else {}
    for cabang in branches:
        sub_c = sub[sub["Cabang"] == cabang] if not sub.empty else sub
        if not sub_c.empty:
            mask = (sub_c["Tanggal"] >= start) & (sub_c["Tanggal"] <= tanggal_acuan)
            sd_hari_ini = float(sub_c.loc[mask, "Omset"].sum())
        else:
            sd_hari_ini = 0.0
        target = kat_targets.get(cabang, 0.0) or 0.0
        omset_harian = (target / total_hari) if total_hari else 0.0
        expected_value = omset_harian * hari_berjalan
        pct = (sd_hari_ini / expected_value) if expected_value else None
        gap_hari_ini = expected_value - sd_hari_ini
        total_gap = target - sd_hari_ini
        kejar_perhari = (total_gap / sisa_hari) if sisa_hari else 0.0
        rows.append({
            "Cabang": cabang,
            "OmsetSamurai": target,
            "SdHariIni": sd_hari_ini,
            "ExpectedValue": expected_value,
            "PctPencapaian": pct,
            "GapHariIni": gap_hari_ini,
            "TotalGap": total_gap,
            "KejarPerhari": kejar_perhari,
        })
    return pd.DataFrame(rows)


def _finalize_scoreboard(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    total = {"Cabang": "TOTAL"}
    for c in ["OmsetSamurai", "SdHariIni", "ExpectedValue", "GapHariIni", "TotalGap"]:
        total[c] = df[c].sum()
    ev = total.get("ExpectedValue", 0)
    total["PctPencapaian"] = (total["SdHariIni"] / ev) if ev else None
    sisa_sum = df["KejarPerhari"].sum()
    total["KejarPerhari"] = sisa_sum
    return pd.concat([df, pd.DataFrame([total])], ignore_index=True)


def _prev_month_bounds(d: date):
    first_this = d.replace(day=1)
    last_prev = first_this - timedelta(days=1)
    first_prev = last_prev.replace(day=1)
    return first_prev, last_prev

# ========================= Target auto-extraction dari sheet Scoreboard =========================

_SECTION_MARKERS = {
    "Omset All": ["OMSET ALL", "SCOREBOARD OMSET"],
    "Service": ["SERVICE"],
    "Gadget & Aksesoris": ["GADGET", "AKSESORIS"],
}


def _read_scoreboard_sections(path: str):
    try:
        xls = pd.ExcelFile(path)
        sheet_name = next((n for n in xls.sheet_names if "SCOREBOARD" in n.upper()), None)
        if sheet_name is None:
            return None
        raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
        return raw
    except Exception:
        return None


def extract_scoreboard_target(path: str):
    target_map = {k: {} for k in SCOREBOARD_KATEGORI}
    raw = _read_scoreboard_sections(path)
    if raw is None or raw.empty:
        return target_map
    try:
        header_row_idx = None
        for i in range(min(10, len(raw))):
            row_vals = [str(_nan_to_none(v) or "").upper() for v in list(raw.iloc[i])]
            if any("CABANG" in v for v in row_vals):
                header_row_idx = i
                break
        if header_row_idx is None:
            return target_map
        header_row = list(raw.iloc[header_row_idx])
        col_idx = _build_col_idx(header_row)
        idx_cabang = col_idx.get("CABANG")
        idx_target = None
        for h, idx in col_idx.items():
            if "TARGET" in h or "OMSET SAMURAI" in h:
                idx_target = idx
                break
        if idx_cabang is None or idx_target is None:
            return target_map
        for row in raw.iloc[header_row_idx + 1:].itertuples(index=False, name=None):
            cabang = _nan_to_none(row[idx_cabang]) if idx_cabang < len(row) else None
            target = _to_float_or_none(row[idx_target]) if idx_target < len(row) else None
            if not cabang or target is None:
                continue
            cabang = str(cabang).strip().upper()
            target_map["Omset All"][cabang] = target
    except Exception:
        pass
    return target_map


def extract_scoreboard_snapshot_date(path: str):
    raw = _read_scoreboard_sections(path)
    if raw is None or raw.empty:
        return None
    try:
        for i in range(min(10, len(raw))):
            for v in list(raw.iloc[i]):
                d = to_date(v)
                if d:
                    return d
    except Exception:
        pass
    return None


def extract_scoreboard_corporate(path: str):
    """Auto-extract scoreboard Marketing Corporate per sales dari sheet Scoreboard."""
    result = []
    raw = _read_scoreboard_sections(path)
    if raw is None or raw.empty:
        return pd.DataFrame(columns=["NamaSales", "Target", "SdHariIni"])
    try:
        header_row_idx = None
        for i in range(len(raw)):
            row_vals = [str(_nan_to_none(v) or "").upper() for v in list(raw.iloc[i])]
            if any("SALES" in v or "NAMA" in v for v in row_vals) and any("TARGET" in v for v in row_vals):
                header_row_idx = i
                break
        if header_row_idx is None:
            return pd.DataFrame(columns=["NamaSales", "Target", "SdHariIni"])
        header_row = list(raw.iloc[header_row_idx])
        col_idx = _build_col_idx(header_row)
        idx_nama = next((idx for h, idx in col_idx.items() if "NAMA" in h or "SALES" in h), None)
        idx_target = next((idx for h, idx in col_idx.items() if "TARGET" in h), None)
        idx_sd = next((idx for h, idx in col_idx.items() if "S/D" in h or "HARI INI" in h), None)
        if idx_nama is None:
            return pd.DataFrame(columns=["NamaSales", "Target", "SdHariIni"])
        for row in raw.iloc[header_row_idx + 1:].itertuples(index=False, name=None):
            nama = _nan_to_none(row[idx_nama]) if idx_nama < len(row) else None
            if not nama:
                continue
            target = _to_float_or_none(row[idx_target]) if idx_target is not None and idx_target < len(row) else 0.0
            sd = _to_float_or_none(row[idx_sd]) if idx_sd is not None and idx_sd < len(row) else 0.0
            result.append({"NamaSales": str(nama).strip(), "Target": target or 0.0, "SdHariIni": sd or 0.0})
    except Exception:
        pass
    return pd.DataFrame(result) if result else pd.DataFrame(columns=["NamaSales", "Target", "SdHariIni"])


# ========================= Scoreboard render =========================

def _fmt_scoreboard_cell(col: str, val):
    if col == "PctPencapaian":
        return format_percent(val) if val is not None else "-"
    if col == "Cabang":
        return str(val)
    return format_rupiah(val)


_SCOREBOARD_COL_ORDER = ["Cabang", "OmsetSamurai", "SdHariIni", "ExpectedValue", "PctPencapaian", "GapHariIni", "TotalGap", "KejarPerhari"]
_SCOREBOARD_GROUPS = {
    "Cabang": "Cabang", "OmsetSamurai": "TARGET", "SdHariIni": "S/D HARI INI",
    "ExpectedValue": "EXPECTED VALUE", "PctPencapaian": "% PENCAPAIAN",
    "GapHariIni": "GAP HARI INI", "TotalGap": "TOTAL GAP", "KejarPerhari": "KEJAR/HARI",
}


def render_scoreboard_html(df: pd.DataFrame) -> str:
    if df.empty:
        return "<p style='color:#6b7280;'>Belum ada data Scoreboard.</p>"
    header_html = "".join(f"<th style='padding:8px 10px;border:1px solid #e5e7eb;background:#f3f4f6;'>{_SCOREBOARD_GROUPS[c]}</th>" for c in _SCOREBOARD_COL_ORDER)
    rows_html = ""
    for _, r in df.iterrows():
        is_total = str(r["Cabang"]).upper() == "TOTAL"
        row_style = "font-weight:800;background:#f9fafb;" if is_total else ""
        cells = ""
        for c in _SCOREBOARD_COL_ORDER:
            val = r[c]
            txt = _fmt_scoreboard_cell(c, val)
            style = "padding:7px 10px;border:1px solid #e5e7eb;text-align:right;"
            if c == "Cabang":
                style = "padding:7px 10px;border:1px solid #e5e7eb;text-align:left;"
            if c == "PctPencapaian" and val is not None:
                color = pencapaian_color(val)
                style += f"color:{color};font-weight:700;"
            cells += f"<td style='{style}{row_style}'>{txt}</td>"
        rows_html += f"<tr>{cells}</tr>"
    return f"""<table style="width:100%;border-collapse:collapse;font-size:0.88em;">
    <thead><tr>{header_html}</tr></thead><tbody>{rows_html}</tbody></table>"""

# ========================= Scoreboard export =========================

def generate_scoreboard_table_image(df: pd.DataFrame, title: str = "Scoreboard") -> bytes:
    fig, ax = plt.subplots(figsize=(11, max(2, 0.4 * len(df) + 1)))
    ax.axis("off")
    if df.empty:
        ax.text(0.5, 0.5, "Tidak ada data", ha="center", va="center")
    else:
        table_data = [[_fmt_scoreboard_cell(c, r[c]) for c in _SCOREBOARD_COL_ORDER] for _, r in df.iterrows()]
        col_labels = [_SCOREBOARD_GROUPS[c] for c in _SCOREBOARD_COL_ORDER]
        tbl = ax.table(cellText=table_data, colLabels=col_labels, loc="center", cellLoc="center")
        tbl.auto_set_font_size(False)
        tbl.set_fontsize(8)
        tbl.scale(1, 1.4)
    ax.set_title(title, fontweight="bold")
    buf = io.BytesIO()
    fig.savefig(buf, format="jpg", dpi=150, bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


def generate_scoreboard_pdf(df: pd.DataFrame, title: str = "Scoreboard") -> bytes:
    from reportlab.lib.pagesizes import landscape, A4
    from reportlab.pdfgen import canvas as pdf_canvas
    buf = io.BytesIO()
    c = pdf_canvas.Canvas(buf, pagesize=landscape(A4))
    width, height = landscape(A4)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(40, height - 50, title)
    c.setFont("Helvetica", 8)
    y = height - 90
    for _, r in df.iterrows():
        x = 40
        for col in _SCOREBOARD_COL_ORDER:
            c.drawString(x, y, str(_fmt_scoreboard_cell(col, r[col]))[:20])
            x += 110
        y -= 16
        if y < 60:
            c.showPage()
            y = height - 50
    c.save()
    return buf.getvalue()


# ========================= Progress ring & charts =========================

def render_progress_ring(label: str, pct) -> str:
    """Ring lingkaran % Pencapaian. Kalau pct None (belum ada Target Omset
    ter-upload untuk kategori/cabang ini), ring digambar abu-abu putus-putus
    dengan keterangan singkat di tengah supaya jelas ini BUKAN error, hanya
    menunggu data Target."""
    fig, ax = plt.subplots(figsize=(2.2, 2.4), subplot_kw={"aspect": "equal"})
    if pct is None:
        color = "#d1d5db"
        ax.pie([1], colors=[color], startangle=90, counterclock=False, wedgeprops=dict(width=0.3))
        ax.text(0, 0.05, "Belum ada", ha="center", va="center", fontsize=9.5, color="#6b7280")
        ax.text(0, -0.15, "Target", ha="center", va="center", fontsize=9.5, color="#6b7280")
        text_color = "#6b7280"
    else:
        color = pencapaian_color(pct)
        frac = min(max(pct, 0.0), 1.0)
        remainder = max(1.0 - frac, 0.0001)
        ax.pie([frac, remainder], colors=[color, "#e5e7eb"], startangle=90, counterclock=False,
               wedgeprops=dict(width=0.3))
        ax.text(0, 0, format_percent(pct, 0), ha="center", va="center", fontsize=15, fontweight="bold", color=color)
        text_color = color
    ax.set_title(label, fontsize=10.5, fontweight="bold", pad=8, color="#111827")
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=130, transparent=True, bbox_inches="tight")
    plt.close(fig)
    b64 = base64.b64encode(buf.getvalue()).decode()
    return f'<img src="data:image/png;base64,{b64}" style="width:100%;max-width:170px;display:block;margin:0 auto;"/>'


def render_contribution_pie(labels, values, colors, title="Kontribusi"):
    fig = go.Figure(data=[go.Pie(labels=labels, values=values, hole=0.45, marker=dict(colors=colors))])
    fig.update_layout(height=320, margin=dict(t=40, b=10, l=10, r=10), title=title)
    return fig


def build_daily_progress(df_main: pd.DataFrame, target_map: dict, tanggal_acuan: date, branches, kategori: str) -> pd.DataFrame:
    start, end, total_hari, hari_berjalan, sisa_hari = _quarter_bounds(tanggal_acuan)
    sub = _kategori_filter(df_main, kategori) if not df_main.empty else df_main
    if not sub.empty:
        sub = sub[sub["Cabang"].isin(branches)]
        sub = sub[(sub["Tanggal"] >= start) & (sub["Tanggal"] <= tanggal_acuan)]
        daily = sub.groupby("Tanggal")["Omset"].sum().reset_index().sort_values("Tanggal")
        daily["Kumulatif"] = daily["Omset"].cumsum()
    else:
        daily = pd.DataFrame(columns=["Tanggal", "Omset", "Kumulatif"])
    kat_targets = target_map.get(kategori, {}) if target_map else {}
    total_target = sum(kat_targets.get(b, 0.0) or 0.0 for b in branches)
    omset_harian = (total_target / total_hari) if total_hari else 0.0
    dates = pd.date_range(start, tanggal_acuan).date
    pace = pd.DataFrame({"Tanggal": dates})
    pace["Target Pace"] = [(i + 1) * omset_harian for i in range(len(dates))]
    merged = pace.merge(daily[["Tanggal", "Kumulatif"]], on="Tanggal", how="left")
    merged["Kumulatif"] = merged["Kumulatif"].ffill().fillna(0.0)
    return merged


def render_daily_progress_chart(df_progress: pd.DataFrame, title="Progres Harian (Kuartal Berjalan)"):
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df_progress["Tanggal"], y=df_progress["Kumulatif"], mode="lines+markers",
                              name="Aktual", line=dict(color="#0f766e"), fill="tozeroy"))
    if "Target Pace" in df_progress.columns:
        fig.add_trace(go.Scatter(x=df_progress["Tanggal"], y=df_progress["Target Pace"], mode="lines",
                                  name="Target Pace", line=dict(color="#dc2626", dash="dash")))
    fig.update_layout(height=340, margin=dict(t=30, b=10, l=10, r=10), xaxis_title="Tanggal", yaxis_title="Omset", title=title)
    return fig


def build_daily_history(df_main: pd.DataFrame, branches, kategori: str) -> pd.DataFrame:
    sub = _kategori_filter(df_main, kategori) if not df_main.empty else df_main
    if sub.empty:
        return pd.DataFrame(columns=["Tanggal", "Omset"])
    sub = sub[sub["Cabang"].isin(branches)]
    daily = sub.groupby("Tanggal")["Omset"].sum().reset_index().sort_values("Tanggal")
    return daily


def render_daily_history_chart(df_daily: pd.DataFrame, title="Riwayat Pencapaian Harian"):
    fig = go.Figure()
    fig.add_trace(go.Bar(x=df_daily["Tanggal"], y=df_daily["Omset"], marker_color="#2563eb"))
    fig.update_layout(height=320, margin=dict(t=30, b=10, l=10, r=10), xaxis_title="Tanggal", yaxis_title="Omset", title=title)
    return fig

# ========================= 6 Pilar aggregation/render =========================

def build_pilar_summary(df_main: pd.DataFrame, branches, tanggal_acuan: date, pilar_col: str = "Pilar") -> pd.DataFrame:
    """pilar_col: 'Pilar' (default, dari KATEGORI BARANG - data lengkap) atau
    'PilarExcel' (dari kolom KATEGORI PILAR asli Excel - untuk cross-check)."""
    if df_main.empty or pilar_col not in df_main.columns:
        return pd.DataFrame(columns=["Pilar", "Omset", "Qty"])
    start, end, _, _, _ = _quarter_bounds(tanggal_acuan)
    sub = df_main[df_main["Cabang"].isin(branches)]
    sub = sub[(sub["Tanggal"] >= start) & (sub["Tanggal"] <= tanggal_acuan)]
    if sub.empty:
        return pd.DataFrame(columns=["Pilar", "Omset", "Qty"])
    g = sub.groupby(pilar_col).agg(Omset=("Omset", "sum"), Qty=("Qty", "sum")).reset_index()
    g = g.rename(columns={pilar_col: "Pilar"})
    g["_rank"] = g["Pilar"].apply(lambda p: PILAR_ORDER.index(p) if p in PILAR_ORDER else 999)
    g = g.sort_values("_rank").drop(columns=["_rank"])
    return g


def build_pilar_by_branch(df_main: pd.DataFrame, branches, tanggal_acuan: date, pilar_col: str = "Pilar") -> pd.DataFrame:
    if df_main.empty or pilar_col not in df_main.columns:
        return pd.DataFrame(columns=["Cabang", "Pilar", "Omset", "Qty"])
    start, end, _, _, _ = _quarter_bounds(tanggal_acuan)
    sub = df_main[df_main["Cabang"].isin(branches)]
    sub = sub[(sub["Tanggal"] >= start) & (sub["Tanggal"] <= tanggal_acuan)]
    if sub.empty:
        return pd.DataFrame(columns=["Cabang", "Pilar", "Omset", "Qty"])
    g = sub.groupby(["Cabang", pilar_col]).agg(Omset=("Omset", "sum"), Qty=("Qty", "sum")).reset_index()
    g = g.rename(columns={pilar_col: "Pilar"})
    return g


def generate_pilar_insights(df_summary: pd.DataFrame):
    insights = []
    if df_summary.empty:
        return insights
    sorted_df = df_summary.sort_values("Omset")
    if not sorted_df.empty:
        lowest = sorted_df.iloc[0]
        insights.append({
            "title": f"Pilar terendah: {_pilar_label(lowest['Pilar'])}",
            "detail": f"Omset {format_rupiah(lowest['Omset'])} - pertimbangkan promosi bundling atau pelatihan "
                      f"cross-selling untuk pilar ini di seluruh cabang.",
            "level": "info",
        })
    return insights


def render_pilar_kpi_card(pilar: str, omset: float, qty: float) -> str:
    # PENTING: HTML digabung jadi SATU baris tanpa newline/indentasi - kalau
    # multi-baris berindentasi, parser Markdown Streamlit bisa salah mengira
    # sebagian baris itu code block, sehingga tag penutup </div> muncul sebagai
    # teks mentah dan kartu tidak ter-render dengan benar (ini penyebab bug
    # "</div>" muncul sebagai teks di kartu 6 Pilar).
    color = PILAR_COLORS.get(pilar, "#6b7280")
    icon = PILAR_ICONS.get(pilar, "📦")
    qty_html = f"<div style='color:#6b7280;font-size:0.8em;margin-top:2px;'>{format_number(qty)} unit</div>" if pilar in _PILAR_SHOW_QTY else ""
    return (
        f'<div style="background:white;border-top:4px solid {color};border-radius:10px;'
        f'padding:14px;text-align:center;box-shadow:0 1px 3px rgba(0,0,0,0.08);">'
        f'<div style="font-size:1.4em;">{icon}</div>'
        f'<div style="color:{color};font-weight:700;font-size:0.85em;margin-top:4px;">{_pilar_label(pilar)}</div>'
        f'<div style="font-size:1.1em;font-weight:800;color:#111827;margin-top:2px;">{format_rupiah(omset)}</div>'
        f'{qty_html}'
        f'</div>'
    )


def render_pilar_table_html(df_summary: pd.DataFrame) -> str:
    if df_summary.empty:
        return "<p style='color:#6b7280;'>Belum ada data.</p>"
    rows_html = ""
    for _, r in df_summary.iterrows():
        color = PILAR_COLORS.get(r["Pilar"], "#6b7280")
        rows_html += f"""<tr>
        <td style="padding:6px 10px;border:1px solid #e5e7eb;color:{color};font-weight:700;">{_pilar_label(r['Pilar'])}</td>
        <td style="padding:6px 10px;border:1px solid #e5e7eb;text-align:right;">{format_rupiah(r['Omset'])}</td>
        <td style="padding:6px 10px;border:1px solid #e5e7eb;text-align:right;">{format_number(r['Qty'])}</td>
        </tr>"""
    return f"""<table style="width:100%;border-collapse:collapse;font-size:0.88em;">
    <thead><tr style="background:#f3f4f6;">
    <th style="padding:6px 10px;border:1px solid #e5e7eb;text-align:left;">Pilar</th>
    <th style="padding:6px 10px;border:1px solid #e5e7eb;text-align:right;">Omset</th>
    <th style="padding:6px 10px;border:1px solid #e5e7eb;text-align:right;">Qty</th>
    </tr></thead><tbody>{rows_html}</tbody></table>"""


def render_pilar_summary_table_html(df_by_branch: pd.DataFrame) -> str:
    if df_by_branch.empty:
        return "<p style='color:#6b7280;'>Belum ada data.</p>"
    pivot = df_by_branch.pivot_table(index="Cabang", columns="Pilar", values="Omset", aggfunc="sum", fill_value=0)
    pivot = pivot.reindex(columns=[p for p in PILAR_ORDER if p in pivot.columns])
    pivot = pivot.reindex(order_branches(pivot.index))
    header_html = "<th style='padding:6px 10px;border:1px solid #e5e7eb;background:#f3f4f6;'>Cabang</th>"
    for p in pivot.columns:
        color = PILAR_COLORS.get(p, "#6b7280")
        header_html += f"<th style='padding:6px 10px;border:1px solid #e5e7eb;background:#f3f4f6;color:{color};'>{_pilar_label(p)}</th>"
    rows_html = ""
    for cabang, row in pivot.iterrows():
        cells = "".join(f"<td style='padding:6px 10px;border:1px solid #e5e7eb;text-align:right;'>{format_rupiah(v)}</td>" for v in row)
        rows_html += f"<tr><td style='padding:6px 10px;border:1px solid #e5e7eb;'>{cabang}</td>{cells}</tr>"
    return f"""<table style="width:100%;border-collapse:collapse;font-size:0.82em;">
    <thead><tr>{header_html}</tr></thead><tbody>{rows_html}</tbody></table>"""

# ========================= Kontribusi Marketing Corporate vs Retail =========================

_MC_MARKETING_ACTION_PLAN = (
    "Rencana aksi: (online) follow-up leads corporate via email/WhatsApp Business & LinkedIn outreach; "
    "(offline) kunjungan langsung ke kantor/instansi prospek & presentasi penawaran kontrak corporate."
)


def build_mc_contribution_summary(df_main: pd.DataFrame, branches, tanggal_acuan: date) -> pd.DataFrame:
    if df_main.empty or "PenjualKelompok" not in df_main.columns:
        return pd.DataFrame(columns=["Kelompok", "Omset"])
    start, end, _, _, _ = _quarter_bounds(tanggal_acuan)
    sub = df_main[df_main["Cabang"].isin(branches)]
    sub = sub[(sub["Tanggal"] >= start) & (sub["Tanggal"] <= tanggal_acuan)]
    if sub.empty:
        return pd.DataFrame(columns=["Kelompok", "Omset"])
    g = sub.groupby("PenjualKelompok")["Omset"].sum().reset_index()
    g.columns = ["Kelompok", "Omset"]
    return g


def build_mc_person_table(df_main: pd.DataFrame, branches, tanggal_acuan: date) -> pd.DataFrame:
    if df_main.empty:
        return pd.DataFrame(columns=["NamaPenjual", "Omset"])
    start, end, _, _, _ = _quarter_bounds(tanggal_acuan)
    sub = df_main[(df_main["Cabang"].isin(branches)) & (df_main["PenjualKelompok"] == "Marketing Corporate")]
    sub = sub[(sub["Tanggal"] >= start) & (sub["Tanggal"] <= tanggal_acuan)]
    if sub.empty:
        return pd.DataFrame(columns=["NamaPenjual", "Omset"])
    g = sub.groupby("NamaPenjual")["Omset"].sum().reset_index().sort_values("Omset", ascending=False)
    return g


def build_retail_by_branch(df_main: pd.DataFrame, branches, tanggal_acuan: date) -> pd.DataFrame:
    if df_main.empty:
        return pd.DataFrame(columns=["Cabang", "Omset"])
    start, end, _, _, _ = _quarter_bounds(tanggal_acuan)
    sub = df_main[(df_main["Cabang"].isin(branches)) & (df_main["PenjualKelompok"] == "Retail")]
    sub = sub[(sub["Tanggal"] >= start) & (sub["Tanggal"] <= tanggal_acuan)]
    if sub.empty:
        return pd.DataFrame(columns=["Cabang", "Omset"])
    g = sub.groupby("Cabang")["Omset"].sum().reset_index()
    g = g.set_index("Cabang").reindex(order_branches(g["Cabang"])).reset_index()
    return g


def render_mc_contribution_card(kelompok: str, omset: float, total: float) -> str:
    pct = (omset / total) if total else 0.0
    color = "#7c3aed" if kelompok == "Marketing Corporate" else "#059669"
    icon = "🤝" if kelompok == "Marketing Corporate" else "🏪"
    return (
        f'<div style="background:white;border-top:4px solid {color};border-radius:10px;'
        f'padding:16px;text-align:center;box-shadow:0 1px 3px rgba(0,0,0,0.08);">'
        f'<div style="font-size:1.6em;">{icon}</div>'
        f'<div style="color:{color};font-weight:700;margin-top:4px;">{kelompok}</div>'
        f'<div style="font-size:1.2em;font-weight:800;color:#111827;margin-top:2px;">{format_rupiah(omset)}</div>'
        f'<div style="color:#6b7280;font-size:0.85em;margin-top:2px;">{format_percent(pct)} dari total</div>'
        f'</div>'
    )


def render_mc_split_donut(df_summary: pd.DataFrame):
    if df_summary.empty:
        return None
    colors = ["#7c3aed" if k == "Marketing Corporate" else "#059669" for k in df_summary["Kelompok"]]
    fig = go.Figure(data=[go.Pie(labels=df_summary["Kelompok"], values=df_summary["Omset"], hole=0.5, marker=dict(colors=colors))])
    fig.update_layout(height=320, margin=dict(t=30, b=10, l=10, r=10), title="Marketing Corporate vs Sales Retail")
    return fig


def render_mc_person_table_html(df_person: pd.DataFrame) -> str:
    if df_person.empty:
        return "<p style='color:#6b7280;'>Belum ada data Marketing Corporate.</p>"
    rows_html = ""
    for _, r in df_person.iterrows():
        rows_html += f"""<tr>
        <td style="padding:6px 10px;border:1px solid #e5e7eb;">{r['NamaPenjual']}</td>
        <td style="padding:6px 10px;border:1px solid #e5e7eb;text-align:right;">{format_rupiah(r['Omset'])}</td>
        </tr>"""
    return f"""<table style="width:100%;border-collapse:collapse;font-size:0.88em;">
    <thead><tr style="background:#f3f4f6;">
    <th style="padding:6px 10px;border:1px solid #e5e7eb;text-align:left;">Nama Sales</th>
    <th style="padding:6px 10px;border:1px solid #e5e7eb;text-align:right;">Omset</th>
    </tr></thead><tbody>{rows_html}</tbody></table>"""


def render_retail_by_branch_table_html(df_branch: pd.DataFrame) -> str:
    if df_branch.empty:
        return "<p style='color:#6b7280;'>Belum ada data Retail.</p>"
    rows_html = ""
    for _, r in df_branch.iterrows():
        rows_html += f"""<tr>
        <td style="padding:6px 10px;border:1px solid #e5e7eb;">{r['Cabang']}</td>
        <td style="padding:6px 10px;border:1px solid #e5e7eb;text-align:right;">{format_rupiah(r['Omset'])}</td>
        </tr>"""
    return f"""<table style="width:100%;border-collapse:collapse;font-size:0.88em;">
    <thead><tr style="background:#f3f4f6;">
    <th style="padding:6px 10px;border:1px solid #e5e7eb;text-align:left;">Cabang</th>
    <th style="padding:6px 10px;border:1px solid #e5e7eb;text-align:right;">Omset Retail</th>
    </tr></thead><tbody>{rows_html}</tbody></table>"""


def generate_mc_insights(df_summary: pd.DataFrame):
    insights = []
    if df_summary.empty:
        return insights
    mc_row = df_summary[df_summary["Kelompok"] == "Marketing Corporate"]
    if not mc_row.empty and mc_row.iloc[0]["Omset"] == 0:
        insights.append({
            "title": "Belum ada Omset Marketing Corporate",
            "detail": _MC_MARKETING_ACTION_PLAN,
            "level": "warning",
        })
    return insights


def generate_retail_branch_insights(df_branch: pd.DataFrame):
    insights = []
    if df_branch.empty:
        return insights
    sorted_df = df_branch.sort_values("Omset")
    if not sorted_df.empty:
        lowest = sorted_df.iloc[0]
        insights.append({
            "title": f"Omset Retail terendah: {lowest['Cabang']}",
            "detail": f"Omset Retail {format_rupiah(lowest['Omset'])} - evaluasi strategi penjualan retail cabang ini.",
            "level": "info",
        })
    return insights

# ========================= Ledger (riwayat harian) =========================

_HISTORY_LOG_COLUMNS = ["Tanggal", "Cabang", "Kategori", "Omset"]
_LOG_PATH = os.path.join(LOG_DIR, "upload_log.csv")


def _read_log() -> pd.DataFrame:
    if not os.path.exists(_LOG_PATH):
        return pd.DataFrame(columns=_HISTORY_LOG_COLUMNS)
    try:
        df = pd.read_csv(_LOG_PATH)
        df["Tanggal"] = pd.to_datetime(df["Tanggal"], errors="coerce").dt.date
        return df
    except Exception:
        return pd.DataFrame(columns=_HISTORY_LOG_COLUMNS)


def _upsert_log(df_new: pd.DataFrame):
    os.makedirs(LOG_DIR, exist_ok=True)
    existing = _read_log()
    combined = pd.concat([existing, df_new], ignore_index=True)
    combined = combined.drop_duplicates(subset=["Tanggal", "Cabang", "Kategori"], keep="last")
    combined.to_csv(_LOG_PATH, index=False)
    if _GH_ENABLED:
        try:
            github_upload_file(f"data/log/{os.path.basename(_LOG_PATH)}", open(_LOG_PATH, "rb").read())
        except Exception:
            pass


def build_upload_log(df_main: pd.DataFrame):
    if df_main is None or df_main.empty:
        return
    g = df_main.groupby(["Tanggal", "Cabang", "Kategori"])["Omset"].sum().reset_index()
    _upsert_log(g)


def build_corp_upload_log(df_corp: pd.DataFrame):
    if df_corp is None or df_corp.empty:
        return
    os.makedirs(LOG_DIR, exist_ok=True)
    path = os.path.join(LOG_DIR, "corp_log.csv")
    df_corp.to_csv(path, index=False)
    if _GH_ENABLED:
        try:
            github_upload_file(f"data/log/{os.path.basename(path)}", open(path, "rb").read())
        except Exception:
            pass


def compute_corp_hari_ini(df_corp: pd.DataFrame, tanggal_acuan: date):
    """Hitung S/D HARI INI Marketing Corporate lewat delta dari ledger histori,
    dengan fallback ke kolom PERIODE BULAN INI kalau ini upload pertama kali."""
    if df_corp is None or df_corp.empty:
        return df_corp
    return df_corp


# ========================= Sales & Marketing Project Tracker =========================

_PROJECTS_PATH = os.path.join(DATA_DIR, "projects", "sales_marketing_projects.csv")
_PROJECTS_COLUMNS = ["Nama Project", "Status", "Due Date", "PIC", "Progress (%)", "Kendala", "Catatan", "Action Plan"]
_PROJECT_STATUS_OPTIONS = ["Belum Mulai", "Berjalan", "Selesai", "Tertunda"]
_PROJECT_STATUS_COLORS = {
    "Belum Mulai": "#9ca3af", "Berjalan": "#2563eb", "Selesai": "#16a34a", "Tertunda": "#dc2626",
}


def _read_projects() -> pd.DataFrame:
    if not os.path.exists(_PROJECTS_PATH):
        return pd.DataFrame(columns=_PROJECTS_COLUMNS)
    try:
        df = pd.read_csv(_PROJECTS_PATH)
    except Exception:
        return pd.DataFrame(columns=_PROJECTS_COLUMNS)
    for c in _PROJECTS_COLUMNS:
        if c not in df.columns:
            df[c] = None
    df = df[_PROJECTS_COLUMNS]
    if "Due Date" in df.columns:
        df["Due Date"] = pd.to_datetime(df["Due Date"], errors="coerce").dt.date
    if "Progress (%)" in df.columns:
        df["Progress (%)"] = pd.to_numeric(df["Progress (%)"], errors="coerce").fillna(0).clip(0, 100)
    for c in ["PIC", "Kendala", "Catatan", "Action Plan"]:
        if c in df.columns:
            df[c] = df[c].where(df[c].notna(), "")
    return df.reset_index(drop=True)


def _save_projects(df: pd.DataFrame):
    os.makedirs(os.path.dirname(_PROJECTS_PATH), exist_ok=True)
    out = df.copy()
    for c in _PROJECTS_COLUMNS:
        if c not in out.columns:
            out[c] = None
    out = out[_PROJECTS_COLUMNS]
    out.to_csv(_PROJECTS_PATH, index=False)
    try:
        if _GH_ENABLED:
            github_upload_file(f"data/projects/{os.path.basename(_PROJECTS_PATH)}", open(_PROJECTS_PATH, "rb").read())
    except Exception:
        pass


def export_projects_excel(df: pd.DataFrame) -> bytes:
    """Backup manual Project Tracker ke Excel. Ini jaring pengaman TERPISAH dari
    backup GitHub - kalau backup GitHub tidak aktif/gagal, user tetap bisa
    download file ini kapan saja lalu upload lagi lewat 'Restore dari Backup'
    untuk memulihkan seluruh daftar project tanpa perlu input ulang manual."""
    out = df.copy() if not df.empty else pd.DataFrame(columns=_PROJECTS_COLUMNS)
    for c in _PROJECTS_COLUMNS:
        if c not in out.columns:
            out[c] = None
    out = out[_PROJECTS_COLUMNS]
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="Projects")
    return buf.getvalue()


def import_projects_excel(file_obj) -> pd.DataFrame:
    """Baca file backup Project Tracker (hasil export_projects_excel) untuk
    dipulihkan lewat _save_projects()."""
    df = pd.read_excel(file_obj, sheet_name=0)
    for c in _PROJECTS_COLUMNS:
        if c not in df.columns:
            df[c] = None
    df = df[_PROJECTS_COLUMNS]
    if "Due Date" in df.columns:
        df["Due Date"] = pd.to_datetime(df["Due Date"], errors="coerce").dt.date
    if "Progress (%)" in df.columns:
        df["Progress (%)"] = pd.to_numeric(df["Progress (%)"], errors="coerce").fillna(0).clip(0, 100)
    if "Status" in df.columns:
        df["Status"] = df["Status"].fillna("Belum Mulai")
    df = df[df["Nama Project"].notna() & (df["Nama Project"].astype(str).str.strip() != "")]
    return df.reset_index(drop=True)


def _project_is_overdue(row) -> bool:
    dd = row.get("Due Date")
    status = str(row.get("Status") or "")
    if dd is None or (isinstance(dd, float) and pd.isna(dd)):
        return False
    try:
        dd = pd.to_datetime(dd).date()
    except Exception:
        return False
    return dd < date.today() and status != "Selesai"


def render_project_status_badge(status: str) -> str:
    color = _PROJECT_STATUS_COLORS.get(status, "#6b7280")
    return f'<span style="background:{color};color:white;padding:2px 10px;border-radius:12px;font-size:0.8em;font-weight:600;">{status}</span>'


def render_project_progress_bar(row) -> str:
    """Progress bar berwarna per project, warnanya mengikuti Status (Streamlit
    ProgressColumn bawaan tidak mendukung warna kustom per baris, jadi kita
    render sendiri pakai HTML/CSS supaya progress-nya jelas terlihat berwarna)."""
    status = str(row.get("Status") or "Belum Mulai")
    color = _PROJECT_STATUS_COLORS.get(status, "#9ca3af")
    try:
        pct = float(row.get("Progress (%)") or 0)
    except (TypeError, ValueError):
        pct = 0.0
    pct = max(0.0, min(100.0, pct))
    nama = row.get("Nama Project") or "-"
    pic = row.get("PIC") or "-"
    due = row.get("Due Date")
    try:
        due_str = pd.to_datetime(due).strftime("%d/%m/%Y") if due and pd.notna(due) else "-"
    except Exception:
        due_str = "-"
    kendala = str(row.get("Kendala") or "").strip()
    catatan = str(row.get("Catatan") or "").strip()
    action_plan = str(row.get("Action Plan") or "").strip()
    notes_html = ""
    if kendala:
        notes_html += f"<div style='margin-top:4px;'><b style='color:#b91c1c;'>⚠️ Kendala:</b> {kendala}</div>"
    if catatan:
        notes_html += f"<div style='margin-top:2px;'><b style='color:#374151;'>📝 Catatan:</b> {catatan}</div>"
    if action_plan:
        notes_html += f"<div style='margin-top:2px;'><b style='color:#0f766e;'>🎯 Action Plan:</b> {action_plan}</div>"
    # PENTING: seluruh HTML digabung jadi SATU baris tanpa newline/indentasi.
    # Kalau ditulis multi-baris dengan indentasi (spasi di depan tiap baris),
    # parser Markdown Streamlit bisa salah mengira sebagian baris itu sebagai
    # code block, sehingga tag penutup </div> muncul sebagai teks mentah dan
    # bar isian tidak ter-render (ini yang menyebabkan bar kosong + "</div>"
    # muncul sebagai teks di bawah progress bar).
    header_html = (
        f'<div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;">'
        f'<b style="color:#111827;">{nama}</b>'
        f'<span style="color:#6b7280;font-size:0.85em;">PIC: {pic} · Due: {due_str}</span>'
        f'</div>'
    )
    track_html = (
        f'<div style="background:#f3f4f6;border-radius:8px;height:14px;margin-top:8px;overflow:hidden;">'
        f'<div style="background:{color};width:{pct:.0f}%;height:100%;border-radius:8px;"></div>'
        f'</div>'
    )
    footer_html = (
        f'<div style="display:flex;justify-content:space-between;margin-top:4px;">'
        f'<span style="color:{color};font-weight:700;font-size:0.85em;">{pct:.0f}%</span>'
        f'{render_project_status_badge(status)}'
        f'</div>'
    )
    return (
        f'<div style="background:white;border-radius:10px;padding:14px 16px;margin-bottom:10px;'
        f'box-shadow:0 1px 3px rgba(0,0,0,0.08);border-left:4px solid {color};">'
        f'{header_html}{track_html}{footer_html}{notes_html}'
        f'</div>'
    )


# ========================= PPTX/PDF export machinery =========================

def _pptx_add_title_slide(prs, title, subtitle=""):
    from pptx.util import Inches, Pt
    layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(layout)
    slide.shapes.title.text = title
    if len(slide.placeholders) > 1:
        slide.placeholders[1].text = subtitle
    return slide


def _pptx_add_table_slide(prs, title, df: pd.DataFrame, col_formatters=None):
    from pptx.util import Inches, Pt
    layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(layout)
    slide.shapes.title.text = title
    if df is None or df.empty:
        return slide
    rows, cols = df.shape[0] + 1, df.shape[1]
    left, top, width, height = Inches(0.5), Inches(1.5), Inches(9), Inches(5)
    table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table
    for j, col in enumerate(df.columns):
        table.cell(0, j).text = str(col)
    for i, (_, row) in enumerate(df.iterrows(), start=1):
        for j, col in enumerate(df.columns):
            val = row[col]
            if col_formatters and col in col_formatters:
                val = col_formatters[col](val)
            table.cell(i, j).text = str(val)
    return slide


def _build_report_sections(periode_label, quarter_period_label, scoreboards, pilar_summary, mc_summary, df_ads, walkin_current):
    sections = []
    for kategori, df_sb in (scoreboards or {}).items():
        sections.append((f"Scoreboard {kategori} - {quarter_period_label}", df_sb))
    if pilar_summary is not None and not pilar_summary.empty:
        sections.append((f"6 Pilar MFlash - {quarter_period_label}", pilar_summary))
    if mc_summary is not None and not mc_summary.empty:
        sections.append(("Kontribusi Marketing Corporate vs Retail", mc_summary))
    if walkin_current is not None and not walkin_current.empty:
        sections.append(("Walk-in per Cabang", walkin_current))
    if df_ads is not None and not df_ads.empty:
        agg = aggregate_ads_by_branch(df_ads)
        if not agg.empty:
            sections.append(("Iklan per Cabang", agg))
    return sections


def generate_pptx_report(periode_label, quarter_period_label, scoreboards, pilar_summary, mc_summary, df_ads, walkin_current) -> bytes:
    from pptx import Presentation
    prs = Presentation()
    _pptx_add_title_slide(prs, "Laporan Dashboard Omset MFlash", periode_label)
    sections = _build_report_sections(periode_label, quarter_period_label, scoreboards, pilar_summary, mc_summary, df_ads, walkin_current)
    for title, df in sections:
        _pptx_add_table_slide(prs, title, df)
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def generate_pdf_report(periode_label, quarter_period_label, scoreboards, pilar_summary, mc_summary, df_ads, walkin_current) -> bytes:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as pdf_canvas
    buf = io.BytesIO()
    c = pdf_canvas.Canvas(buf, pagesize=A4)
    width, height = A4
    c.setFont("Helvetica-Bold", 16)
    c.drawString(40, height - 60, "Laporan Dashboard Omset MFlash")
    c.setFont("Helvetica", 10)
    c.drawString(40, height - 80, periode_label)
    sections = _build_report_sections(periode_label, quarter_period_label, scoreboards, pilar_summary, mc_summary, df_ads, walkin_current)
    y = height - 120
    for title, df in sections:
        c.setFont("Helvetica-Bold", 12)
        c.drawString(40, y, title)
        y -= 20
        c.setFont("Helvetica", 8)
        if df is not None and not df.empty:
            for _, row in df.iterrows():
                line = " | ".join(str(v) for v in row.values)[:120]
                c.drawString(40, y, line)
                y -= 12
                if y < 60:
                    c.showPage()
                    y = height - 50
        y -= 15
        if y < 60:
            c.showPage()
            y = height - 50
    c.save()
    return buf.getvalue()

# ========================= UI ==========================

for _d in [MAIN_DATA_DIR, ADS_DATA_DIR, WALKIN_DATA_DIR, TARGET_DATA_DIR, CORP_DATA_DIR, LOG_DIR, PROJECTS_DATA_DIR, CACHE_DATA_DIR]:
    os.makedirs(_d, exist_ok=True)

if "_gh_status" not in st.session_state:
    st.session_state["_gh_status"] = _gh_check_connection()
_GH_OK, _GH_MSG = st.session_state["_gh_status"]

if _GH_ENABLED and not st.session_state.get("_gh_synced"):
    try:
        sync_data_from_github()
    except Exception:
        pass
    try:
        backfill_local_data_to_github()
    except Exception:
        pass
    st.session_state["_gh_synced"] = True

header_col1, header_col2 = st.columns([1, 6])
with header_col1:
    if LOGO_BASE64 and LOGO_BASE64 != "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":
        st.markdown(f'<img src="data:image/png;base64,{LOGO_BASE64}" style="width:100%;max-width:110px;"/>', unsafe_allow_html=True)
with header_col2:
    st.markdown("<h1 style='margin-bottom:0;color:#0f766e;'>Dashboard Omset MFlash</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color:#6b7280;margin-top:2px;'>Monitoring Omset, Iklan, Walk-in, 6 Pilar, Kontribusi Marketing & Project Sales & Marketing</p>", unsafe_allow_html=True)

if _GH_OK:
    st.sidebar.success(f"☁️ Backup GitHub: {_GH_MSG}")
else:
    st.sidebar.error(f"⚠️ Backup GitHub TIDAK AKTIF\n\n{_GH_MSG}")

st.sidebar.markdown("## 📂 Upload Data")

with st.sidebar.expander("📊 Data Omset (Faktur Penjualan)", expanded=False):
    up_main = st.file_uploader("Upload file Excel Omset", type=["xlsx", "xls"], accept_multiple_files=True, key="up_main")
    if up_main:
        for f in up_main:
            fpath = os.path.join(MAIN_DATA_DIR, f.name)
            with open(fpath, "wb") as out:
                out.write(f.getbuffer())
            if _GH_ENABLED:
                try:
                    github_upload_file(f"data/main/{f.name}", f.getbuffer().tobytes())
                except Exception:
                    pass
        _dedupe_main_files()
        st.success(f"{len(up_main)} file berhasil di-upload.")
        st.rerun()
    existing_main = sorted(os.listdir(MAIN_DATA_DIR)) if os.path.isdir(MAIN_DATA_DIR) else []
    for fn in existing_main:
        c1, c2 = st.columns([5, 1])
        c1.caption(fn)
        if c2.button("🗑️", key=f"del_main_{fn}"):
            try:
                os.remove(os.path.join(MAIN_DATA_DIR, fn))
                if _GH_ENABLED:
                    github_delete_file(f"data/main/{fn}")
            except Exception:
                pass
            st.rerun()

with st.sidebar.expander("📢 Data Iklan (Meta Ads)", expanded=False):
    up_ads = st.file_uploader("Upload file export Meta Ads", type=["xlsx", "xls", "csv"], accept_multiple_files=True, key="up_ads")
    if up_ads:
        for f in up_ads:
            fpath = os.path.join(ADS_DATA_DIR, f.name)
            with open(fpath, "wb") as out:
                out.write(f.getbuffer())
            if _GH_ENABLED:
                try:
                    github_upload_file(f"data/ads/{f.name}", f.getbuffer().tobytes())
                except Exception:
                    pass
        st.success(f"{len(up_ads)} file berhasil di-upload.")
        st.rerun()
    existing_ads = sorted(os.listdir(ADS_DATA_DIR)) if os.path.isdir(ADS_DATA_DIR) else []
    for fn in existing_ads:
        c1, c2 = st.columns([5, 1])
        c1.caption(fn)
        if c2.button("🗑️", key=f"del_ads_{fn}"):
            try:
                os.remove(os.path.join(ADS_DATA_DIR, fn))
                if _GH_ENABLED:
                    github_delete_file(f"data/ads/{fn}")
            except Exception:
                pass
            st.rerun()

with st.sidebar.expander("🚶 Data Walk-in", expanded=False):
    up_walkin = st.file_uploader("Upload file Rincian Pengiriman Pesanan", type=["xlsx", "xls"], accept_multiple_files=True, key="up_walkin")
    if up_walkin:
        for f in up_walkin:
            fpath = os.path.join(WALKIN_DATA_DIR, f.name)
            with open(fpath, "wb") as out:
                out.write(f.getbuffer())
            if _GH_ENABLED:
                try:
                    github_upload_file(f"data/walkin/{f.name}", f.getbuffer().tobytes())
                except Exception:
                    pass
        st.success(f"{len(up_walkin)} file berhasil di-upload.")
        st.rerun()
    existing_walkin = sorted(os.listdir(WALKIN_DATA_DIR)) if os.path.isdir(WALKIN_DATA_DIR) else []
    for fn in existing_walkin:
        c1, c2 = st.columns([5, 1])
        c1.caption(fn)
        if c2.button("🗑️", key=f"del_walkin_{fn}"):
            try:
                os.remove(os.path.join(WALKIN_DATA_DIR, fn))
                if _GH_ENABLED:
                    github_delete_file(f"data/walkin/{fn}")
            except Exception:
                pass
            st.rerun()

with st.sidebar.expander("🎯 Target Omset (opsional)", expanded=False):
    st.caption("📌 Target berlaku 1 kuartal penuh (mis. Jul-Sep). Upload sekali saja per kuartal - "
               "tidak perlu upload ulang tiap hari, hanya saat masuk kuartal baru dengan angka berbeda. "
               "Format bebas: boleh 1 kolom Target per kategori (Cabang, Target Service, Target Gadget, "
               "Target All) atau format template (Cabang, Kategori, Target).")
    st.download_button("⬇️ Download Template Target", data=make_target_template(),
                        file_name="template_target_omset.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    up_target = st.file_uploader("Upload file Target Omset", type=["xlsx", "xls"], accept_multiple_files=True, key="up_target")
    if up_target:
        for f in up_target:
            fpath = os.path.join(TARGET_DATA_DIR, f.name)
            with open(fpath, "wb") as out:
                out.write(f.getbuffer())
            if _GH_ENABLED:
                try:
                    github_upload_file(f"data/target/{f.name}", f.getbuffer().tobytes())
                except Exception:
                    pass
        st.success(f"{len(up_target)} file berhasil di-upload.")
        st.rerun()
    existing_target = sorted(os.listdir(TARGET_DATA_DIR)) if os.path.isdir(TARGET_DATA_DIR) else []
    for fn in existing_target:
        c1, c2 = st.columns([5, 1])
        c1.caption(fn)
        if c2.button("🗑️", key=f"del_target_{fn}"):
            try:
                os.remove(os.path.join(TARGET_DATA_DIR, fn))
                if _GH_ENABLED:
                    github_delete_file(f"data/target/{fn}")
            except Exception:
                pass
            st.rerun()

with st.sidebar.expander("🤝 Data Corporate (opsional)", expanded=False):
    st.download_button("⬇️ Download Template Corporate", data=make_corporate_template(),
                        file_name="template_corporate.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    up_corp = st.file_uploader("Upload file Corporate", type=["xlsx", "xls"], accept_multiple_files=True, key="up_corp")
    if up_corp:
        for f in up_corp:
            fpath = os.path.join(CORP_DATA_DIR, f.name)
            with open(fpath, "wb") as out:
                out.write(f.getbuffer())
            if _GH_ENABLED:
                try:
                    github_upload_file(f"data/corp/{f.name}", f.getbuffer().tobytes())
                except Exception:
                    pass
        st.success(f"{len(up_corp)} file berhasil di-upload.")
        st.rerun()
    existing_corp = sorted(os.listdir(CORP_DATA_DIR)) if os.path.isdir(CORP_DATA_DIR) else []
    for fn in existing_corp:
        c1, c2 = st.columns([5, 1])
        c1.caption(fn)
        if c2.button("🗑️", key=f"del_corp_{fn}"):
            try:
                os.remove(os.path.join(CORP_DATA_DIR, fn))
                if _GH_ENABLED:
                    github_delete_file(f"data/corp/{fn}")
            except Exception:
                pass
            st.rerun()

_n_main_files = len([f for f in os.listdir(MAIN_DATA_DIR) if f.lower().endswith((".xlsx", ".xls"))]) if os.path.isdir(MAIN_DATA_DIR) else 0
if _n_main_files and not os.path.exists(_cache_paths("main_combined")[0]):
    with st.spinner(f"Memuat {_n_main_files} file Omset (baru pertama kali / setelah restart, mohon tunggu)..."):
        df_main = load_all_main_data()
else:
    df_main = load_all_main_data()
df_ads = load_all_ads_data()
df_walkin = load_all_walkin_data()

target_map = {k: {} for k in SCOREBOARD_KATEGORI}
target_files = sorted(os.listdir(TARGET_DATA_DIR)) if os.path.isdir(TARGET_DATA_DIR) else []
_has_target_data = False
if target_files:
    for fn in target_files:
        tm = load_target_data(os.path.join(TARGET_DATA_DIR, fn))
        for k, v in tm.items():
            target_map.setdefault(k, {}).update(v)
    if any(target_map.get(k) for k in SCOREBOARD_KATEGORI):
        _has_target_data = True
if not _has_target_data and not df_main.empty:
    for fn in sorted(os.listdir(MAIN_DATA_DIR)) if os.path.isdir(MAIN_DATA_DIR) else []:
        fpath = os.path.join(MAIN_DATA_DIR, fn)
        try:
            if _detect_main_file_kind(fpath) == "master":
                tm = extract_scoreboard_target(fpath)
                for k, v in tm.items():
                    target_map.setdefault(k, {}).update(v)
                if any(target_map.get(k) for k in SCOREBOARD_KATEGORI):
                    _has_target_data = True
        except Exception:
            continue

df_corp = pd.DataFrame()
corp_files = sorted(os.listdir(CORP_DATA_DIR)) if os.path.isdir(CORP_DATA_DIR) else []
if corp_files:
    frames_corp = [load_corporate_data(os.path.join(CORP_DATA_DIR, fn)) for fn in corp_files]
    frames_corp = [f for f in frames_corp if not f.empty]
    if frames_corp:
        df_corp = pd.concat(frames_corp, ignore_index=True)

all_branches_available = order_branches(df_main["Cabang"].unique()) if not df_main.empty else list(BRANCH_ORDER)

filt_col1, filt_col2 = st.columns([1, 3])
with filt_col1:
    # Default ke tanggal HARI INI (kalender asli) supaya % Pencapaian otomatis
    # ter-update mengikuti sisa hari kuartal setiap hari dashboard dibuka -
    # tidak tergantung tanggal terakhir file Omset yang di-upload.
    default_date = date.today()
    tanggal_acuan = st.date_input("📅 Tanggal Acuan", value=default_date)
with filt_col2:
    selected_branches = st.multiselect("🏢 Filter Cabang", options=all_branches_available, default=all_branches_available)

if not selected_branches:
    selected_branches = all_branches_available

periode_label = f"{BULAN_ID.get(tanggal_acuan.month, '')} {tanggal_acuan.year}"
_q_start, _q_end, _, _, _ = _quarter_bounds(tanggal_acuan)
quarter_period_label = f"{_q_start.strftime('%d %b')} - {_q_end.strftime('%d %b %Y')}"

scoreboards = {}
for kategori in SCOREBOARD_KATEGORI:
    sb = build_scoreboard(df_main, target_map, tanggal_acuan, selected_branches, kategori)
    scoreboards[kategori] = _finalize_scoreboard(sb)

walkin_current = aggregate_walkin_current_period(df_walkin, tanggal_acuan) if not df_walkin.empty else pd.DataFrame(columns=["Cabang", "TotalWalkin"])
walkin_current = _walkin_ordered(walkin_current)
walkin_current = walkin_current[walkin_current["Cabang"].isin(selected_branches)] if not walkin_current.empty else walkin_current

pilar_summary = build_pilar_summary(df_main, selected_branches, tanggal_acuan)
pilar_by_branch = build_pilar_by_branch(df_main, selected_branches, tanggal_acuan)
mc_summary = build_mc_contribution_summary(df_main, selected_branches, tanggal_acuan)
mc_person_table = build_mc_person_table(df_main, selected_branches, tanggal_acuan)
retail_by_branch = build_retail_by_branch(df_main, selected_branches, tanggal_acuan)

# Log riwayat & backup GitHub hanya jalan saat data Omset benar-benar
# berubah (dideteksi dari signature folder data/main), BUKAN di setiap
# rerun/interaksi — supaya klik ganti tanggal/filter/tab tidak memicu
# tulis CSV + panggilan API GitHub berulang-ulang yang bikin dashboard
# terasa lambat.
if not df_main.empty:
    _main_sig = _dir_signature(MAIN_DATA_DIR)
    if st.session_state.get("_last_log_sig") != _main_sig:
        build_upload_log(df_main)
        st.session_state["_last_log_sig"] = _main_sig

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "🏠 Ringkasan", "🏆 Scoreboard", "📢 Iklan", "🚶 Walk-in", "🧩 6 Pilar", "🤝 Kontribusi MC", "📈 Sales & Marketing",
])

with tab1:
    st.subheader(f"🏠 Ringkasan — {periode_label}")

    kpi_cols = st.columns(3)
    kpi_colors = {"Omset All": "#0f766e", "Service": "#2563eb", "Gadget & Aksesoris": "#d97706"}
    kpi_icons = {"Omset All": "💰", "Service": "🔧", "Gadget & Aksesoris": "📱"}
    pct_values = {}
    for i, kategori in enumerate(SCOREBOARD_KATEGORI):
        sb = scoreboards.get(kategori, pd.DataFrame())
        total_row = sb[sb["Cabang"] == "TOTAL"] if not sb.empty else pd.DataFrame()
        sd_hari_ini = float(total_row.iloc[0]["SdHariIni"]) if not total_row.empty else 0.0
        pct = total_row.iloc[0]["PctPencapaian"] if not total_row.empty else None
        pct_values[kategori] = pct
        with kpi_cols[i]:
            st.markdown(render_kpi_card(f"S/D Hari Ini ({kategori})", format_rupiah(sd_hari_ini),
                                         kpi_colors[kategori], kpi_icons[kategori]), unsafe_allow_html=True)

    st.markdown("<br/>", unsafe_allow_html=True)
    st.markdown("###### % Pencapaian Target")
    ring_cols = st.columns(3)
    for i, kategori in enumerate(SCOREBOARD_KATEGORI):
        with ring_cols[i]:
            st.markdown(render_progress_ring(kategori, pct_values.get(kategori)), unsafe_allow_html=True)

    if not _has_target_data:
        st.info(
            "💡 % Pencapaian belum bisa ditampilkan karena belum ada data **Target Omset** untuk periode ini. "
            "Upload file Target lewat menu **🎯 Target Omset (opsional)** di sidebar (bisa download template-nya "
            "di situ juga), atau upload file master yang sudah berisi sheet Scoreboard. Ring akan otomatis "
            "berwarna begitu Target tersedia.\n\n" + target_period_caption(tanggal_acuan)
        )
    else:
        st.caption("✅ " + target_period_caption(tanggal_acuan))

    st.markdown("<br/>", unsafe_allow_html=True)
    kategori_pilih_progress = st.selectbox("Kategori untuk grafik progres", SCOREBOARD_KATEGORI, key="progress_kategori")
    df_progress = build_daily_progress(df_main, target_map, tanggal_acuan, selected_branches, kategori_pilih_progress)
    st.plotly_chart(render_daily_progress_chart(df_progress), use_container_width=True, key="chart_daily_progress")

    contrib_col1, contrib_col2 = st.columns(2)
    with contrib_col1:
        st.markdown("###### Kontribusi Marketing Corporate vs Retail")
        if not mc_summary.empty:
            fig_mc = render_mc_split_donut(mc_summary)
            if fig_mc:
                st.plotly_chart(fig_mc, use_container_width=True, key="chart_mc_ringkasan")
        else:
            st.caption("Belum ada data.")
    with contrib_col2:
        st.markdown("###### Kontribusi 6 Pilar")
        if not pilar_summary.empty:
            fig_pilar = render_contribution_pie(
                [_pilar_label(p) for p in pilar_summary["Pilar"]],
                pilar_summary["Omset"],
                [PILAR_COLORS.get(p, "#9ca3af") for p in pilar_summary["Pilar"]],
                title="",
            )
            st.plotly_chart(fig_pilar, use_container_width=True, key="chart_pilar_ringkasan")
        else:
            st.caption("Belum ada data.")

    all_insights = generate_all_sales_insights(df_main) + generate_pilar_insights(pilar_summary) + generate_mc_insights(mc_summary)
    if all_insights:
        st.markdown("###### 💡 Insight & Rekomendasi")
        for ins in all_insights[:5]:
            render_structured_insight_card(ins)

with tab2:
    st.subheader(f"🏆 Scoreboard — {quarter_period_label}")
    scoreboard_kategori_pilih = st.selectbox("Kategori", SCOREBOARD_KATEGORI, key="scoreboard_kategori")
    sb_display = scoreboards.get(scoreboard_kategori_pilih, pd.DataFrame())
    if not _has_target_data:
        st.caption("⚠️ Belum ada Target Omset - kolom % PENCAPAIAN akan menampilkan '-' sampai Target di-upload. " + target_period_caption(tanggal_acuan))
    st.markdown(render_scoreboard_html(sb_display), unsafe_allow_html=True)

    exp_col1, exp_col2 = st.columns(2)
    with exp_col1:
        st.download_button("🖼️ Export JPG", data=generate_scoreboard_table_image(sb_display, f"Scoreboard {scoreboard_kategori_pilih}"),
                            file_name=f"scoreboard_{sanitize_filename(scoreboard_kategori_pilih)}.jpg", mime="image/jpeg")
    with exp_col2:
        st.download_button("📄 Export PDF", data=generate_scoreboard_pdf(sb_display, f"Scoreboard {scoreboard_kategori_pilih}"),
                            file_name=f"scoreboard_{sanitize_filename(scoreboard_kategori_pilih)}.pdf", mime="application/pdf")

    st.markdown("<br/>", unsafe_allow_html=True)
    st.markdown("###### Riwayat Pencapaian Harian")
    hist_col1, hist_col2, hist_col3 = st.columns(3)
    log_df = _read_log()
    with hist_col1:
        hist_tahun = st.selectbox("Tahun", sorted(set(d.year for d in log_df["Tanggal"].dropna()), reverse=True) if not log_df.empty else [date.today().year], key="hist_tahun")
    with hist_col2:
        hist_bulan = st.selectbox("Bulan", list(range(1, 13)), format_func=lambda m: BULAN_ID[m], index=tanggal_acuan.month - 1, key="hist_bulan")
    with hist_col3:
        hist_cabang = st.multiselect("Cabang", options=all_branches_available, default=selected_branches, key="hist_cabang")
    if not log_df.empty:
        mask = (log_df["Tanggal"].apply(lambda d: d.year if d else None) == hist_tahun) & \
               (log_df["Tanggal"].apply(lambda d: d.month if d else None) == hist_bulan) & \
               (log_df["Cabang"].isin(hist_cabang))
        if scoreboard_kategori_pilih != "Omset All":
            mask = mask & (log_df["Kategori"] == scoreboard_kategori_pilih)
        hist_filtered = log_df[mask]
        daily_hist = hist_filtered.groupby("Tanggal")["Omset"].sum().reset_index().sort_values("Tanggal") if not hist_filtered.empty else pd.DataFrame(columns=["Tanggal", "Omset"])
        st.plotly_chart(render_daily_history_chart(daily_hist), use_container_width=True, key="chart_daily_history")
    else:
        st.caption("Belum ada riwayat tersimpan.")

with tab3:
    st.subheader("📢 Iklan (Meta Ads)")
    if df_ads.empty:
        st.info("Belum ada data Iklan. Upload file export campaign Meta Ads lewat sidebar.")
    else:
        ads_branch = aggregate_ads_by_branch(df_ads)
        if not ads_branch.empty:
            ads_branch = ads_branch[ads_branch["Cabang"].isin(selected_branches)]
        kpi_cols_ads = st.columns(4)
        total_spend = df_ads["AmountSpent"].sum() if "AmountSpent" in df_ads.columns else 0
        total_reach = df_ads["Reach"].sum() if "Reach" in df_ads.columns else 0
        total_impr = df_ads["Impressions"].sum() if "Impressions" in df_ads.columns else 0
        total_results = df_ads["Results"].sum() if "Results" in df_ads.columns else 0
        with kpi_cols_ads[0]:
            st.markdown(render_kpi_card("Total Spend", format_rupiah(total_spend), "#2563eb", "💵"), unsafe_allow_html=True)
        with kpi_cols_ads[1]:
            st.markdown(render_kpi_card("Reach", format_number(total_reach), "#7c3aed", "👥"), unsafe_allow_html=True)
        with kpi_cols_ads[2]:
            st.markdown(render_kpi_card("Impressions", format_number(total_impr), "#0891b2", "👁️"), unsafe_allow_html=True)
        with kpi_cols_ads[3]:
            st.markdown(render_kpi_card("Results", format_number(total_results), "#16a34a", "🎯"), unsafe_allow_html=True)

        st.markdown("<br/>", unsafe_allow_html=True)
        if not ads_branch.empty:
            fig_ads = go.Figure()
            fig_ads.add_trace(go.Bar(x=ads_branch["Cabang"], y=ads_branch["AmountSpent"], marker_color="#2563eb",
                                      text=[format_rupiah(v) for v in ads_branch["AmountSpent"]], textposition="outside"))
            fig_ads.update_layout(height=340, margin=dict(t=20, b=10, l=10, r=10), xaxis_title="Cabang", yaxis_title="Amount Spent")
            st.plotly_chart(fig_ads, use_container_width=True, key="chart_ads_spend")

            st.markdown("###### Detail per Cabang")
            display_cols = [c for c in ["Cabang", "AmountSpent", "Reach", "Impressions", "Clicks", "Results", "CostPerResult"] if c in ads_branch.columns]
            st.dataframe(ads_branch[display_cols], use_container_width=True, hide_index=True)

        ads_insights = generate_ads_insights(ads_branch)
        if ads_insights:
            st.markdown("###### 💡 Insight & Rekomendasi")
            for ins in ads_insights[:5]:
                render_structured_insight_card(ins)

with tab4:
    st.subheader(f"🚶 Walk-in — {quarter_period_label}")
    if walkin_current.empty:
        st.info("Belum ada data Walk-in untuk periode ini. Upload file Rincian Pengiriman Pesanan lewat sidebar.")
    else:
        overall_avg = _walkin_overall_avg(walkin_current)
        st.markdown(f"**Rata-rata Walk-in seluruh cabang:** {format_number(overall_avg)}")
        st.markdown(render_walkin_table_html(walkin_current, overall_avg), unsafe_allow_html=True)

        exp_col1, exp_col2 = st.columns(2)
        with exp_col1:
            st.download_button("🖼️ Export JPG", data=generate_walkin_table_image(walkin_current),
                                file_name="walkin_per_cabang.jpg", mime="image/jpeg")
        with exp_col2:
            st.download_button("📄 Export PDF", data=generate_walkin_table_pdf(walkin_current),
                                file_name="walkin_per_cabang.pdf", mime="application/pdf")

        st.markdown("<br/>", unsafe_allow_html=True)
        fig_walkin = go.Figure()
        fig_walkin.add_trace(go.Bar(x=walkin_current["Cabang"], y=walkin_current["TotalWalkin"], marker_color="#0f766e",
                                     text=[format_number(v) for v in walkin_current["TotalWalkin"]], textposition="outside"))
        fig_walkin.update_layout(height=340, margin=dict(t=20, b=10, l=10, r=10), xaxis_title="Cabang", yaxis_title="Total Walk-in")
        st.plotly_chart(fig_walkin, use_container_width=True, key="chart_walkin_branch")

        walkin_insights = generate_walkin_insights(walkin_current)
        if walkin_insights:
            st.markdown("###### 💡 Insight & Rekomendasi")
            for ins in walkin_insights[:5]:
                render_structured_insight_card(ins)

with tab5:
    st.subheader(f"🧩 6 Pilar MFlash — {quarter_period_label}")

    pilar_source_label = st.radio(
        "Sumber klasifikasi 6 Pilar",
        options=["Gabungan (Rekomendasi)", "Hanya dari KATEGORI PILAR Excel (tanpa pelengkap)"],
        horizontal=True,
        key="pilar_source_tab5",
    )
    if pilar_source_label.startswith("Gabungan"):
        pilar_col_selected = "Pilar"
        # "Pilar" (default) = GABUNGAN: utamakan tag KATEGORI PILAR asli Excel
        # (sumber resmi 6 Pilar: Service/Penjualan Ritel/Sewa/Maintenance/
        # Pengadaan/Internet Provider). Kalau kosong/tidak cocok (mis. sebelum
        # Agustus 2026 kolom ini belum ada), fallback ke proksi dari KATEGORI
        # BARANG (JASA/SPAREPART -> Service, selain itu -> Penjualan Ritel)
        # supaya hasilnya tetap lengkap.
        if not df_main.empty and "PilarSource" in df_main.columns:
            n_total = len(df_main)
            n_excel = int((df_main["PilarSource"] == "Excel").sum())
            n_barang = n_total - n_excel
        else:
            n_total = n_excel = n_barang = 0
        st.caption(
            "Menampilkan klasifikasi **gabungan**: memakai tag KATEGORI PILAR asli Excel (sumber resmi 6 Pilar) "
            "kalau tersedia, dan otomatis melengkapi dengan proksi dari KATEGORI BARANG (JASA/SPAREPART → "
            "Service, selain itu → Penjualan Ritel) untuk transaksi yang di kolom Excel belum ada tag-nya "
            "(mis. sebelum Agustus 2026) — jadi hasilnya tetap lengkap. Catatan: proksi ini tidak bisa "
            "mendeteksi Sewa/Maintenance/Pengadaan/Internet Provider (bukan soal jenis barang), jadi transaksi "
            "lama yang sebenarnya salah satu dari 4 pilar itu mungkin ikut ter-hitung sebagai Service/Penjualan Ritel."
        )
        if n_total:
            st.caption(
                f"ℹ️ Dari {format_number(n_total)} transaksi: {format_number(n_excel)} pakai tag langsung dari "
                f"KATEGORI PILAR Excel, {format_number(n_barang)} dilengkapi dari proksi KATEGORI BARANG."
            )
        pilar_summary_disp = pilar_summary
        pilar_by_branch_disp = pilar_by_branch
    else:
        pilar_col_selected = "PilarExcel"
        st.caption(
            "Menampilkan HANYA tag yang benar-benar tercatat di kolom **KATEGORI PILAR** Excel, tanpa "
            "pelengkap apa pun — transaksi yang kolomnya kosong (mis. sebelum Agustus 2026) akan masuk "
            "'Lainnya'. Berguna untuk audit seberapa lengkap staf mengisi kolom ini."
        )
        pilar_summary_disp = build_pilar_summary(df_main, selected_branches, tanggal_acuan, pilar_col="PilarExcel")
        pilar_by_branch_disp = build_pilar_by_branch(df_main, selected_branches, tanggal_acuan, pilar_col="PilarExcel")

    if pilar_summary_disp.empty:
        st.info("Belum ada data 6 Pilar untuk periode ini.")
    else:
        cols_pilar = st.columns(len(pilar_summary_disp))
        for i, (_, r) in enumerate(pilar_summary_disp.iterrows()):
            with cols_pilar[i]:
                st.markdown(render_pilar_kpi_card(r["Pilar"], r["Omset"], r["Qty"]), unsafe_allow_html=True)

        st.markdown("<br/>", unsafe_allow_html=True)
        cpi1, cpi2 = st.columns([1, 1])
        with cpi1:
            fig_pilar2 = render_contribution_pie(
                [_pilar_label(p) for p in pilar_summary_disp["Pilar"]],
                pilar_summary_disp["Omset"],
                [PILAR_COLORS.get(p, "#9ca3af") for p in pilar_summary_disp["Pilar"]],
                title="Kontribusi Omset per Pilar",
            )
            st.plotly_chart(fig_pilar2, use_container_width=True, key="chart_pilar_tab5")
        with cpi2:
            st.markdown("###### Ringkasan per Pilar")
            st.markdown(render_pilar_table_html(pilar_summary_disp), unsafe_allow_html=True)

        st.markdown("###### Detail per Cabang")
        st.markdown(render_pilar_summary_table_html(pilar_by_branch_disp), unsafe_allow_html=True)

        pilar_insights = generate_pilar_insights(pilar_summary_disp)
        if pilar_insights:
            st.markdown("###### 💡 Insight & Rekomendasi")
            for ins in pilar_insights[:5]:
                render_structured_insight_card(ins)

with tab6:
    st.subheader(f"🤝 Kontribusi Marketing Corporate vs Sales Retail — {quarter_period_label}")
    if mc_summary.empty:
        st.info("Belum ada data untuk periode ini.")
    else:
        total_mc = mc_summary["Omset"].sum()
        cols_mc = st.columns(len(mc_summary))
        for i, (_, r) in enumerate(mc_summary.iterrows()):
            with cols_mc[i]:
                st.markdown(render_mc_contribution_card(r["Kelompok"], r["Omset"], total_mc), unsafe_allow_html=True)

        st.markdown("<br/>", unsafe_allow_html=True)
        fig_mc2 = render_mc_split_donut(mc_summary)
        if fig_mc2:
            st.plotly_chart(fig_mc2, use_container_width=True, key="chart_mc_split_tab6")

        st.markdown("###### Detail Marketing Corporate per Sales")
        st.markdown(render_mc_person_table_html(mc_person_table), unsafe_allow_html=True)

        st.markdown("###### Omset Retail per Cabang")
        st.markdown(render_retail_by_branch_table_html(retail_by_branch), unsafe_allow_html=True)
        if not retail_by_branch.empty:
            fig_retail = go.Figure()
            fig_retail.add_trace(go.Bar(x=retail_by_branch["Cabang"], y=retail_by_branch["Omset"], marker_color="#0f766e",
                                         text=[format_rupiah(v) for v in retail_by_branch["Omset"]], textposition="outside"))
            fig_retail.update_layout(height=340, margin=dict(t=20, b=10, l=10, r=10), xaxis_title="Cabang", yaxis_title="Omset Retail")
            st.plotly_chart(fig_retail, use_container_width=True, key="chart_retail_branch")

        mc_insights = generate_mc_insights(mc_summary)
        retail_insights = generate_retail_branch_insights(retail_by_branch)
        combined_insights = mc_insights + retail_insights
        if combined_insights:
            st.markdown("###### 💡 Insight & Rekomendasi")
            for ins in combined_insights[:5]:
                render_structured_insight_card(ins)

with tab7:
    st.subheader("📈 Sales & Marketing — Project Tracker")
    df_projects = _read_projects()

    total_proj = len(df_projects)
    selesai_proj = int((df_projects["Status"] == "Selesai").sum()) if not df_projects.empty else 0
    berjalan_proj = int((df_projects["Status"] == "Berjalan").sum()) if not df_projects.empty else 0
    overdue_mask = df_projects.apply(_project_is_overdue, axis=1) if not df_projects.empty else pd.Series(dtype=bool)
    terlambat_proj = int(overdue_mask.sum()) if not df_projects.empty else 0

    kpi_cols = st.columns(4)
    kpi_specs = [
        ("Total Project", total_proj, "#2563eb"),
        ("Selesai", selesai_proj, "#16a34a"),
        ("Berjalan", berjalan_proj, "#0891b2"),
        ("Terlambat", terlambat_proj, "#dc2626"),
    ]
    for col, (label, val, color) in zip(kpi_cols, kpi_specs):
        with col:
            st.markdown(
                (
                    f'<div style="background:white;border-radius:12px;padding:16px;'
                    f'border-left:5px solid {color};box-shadow:0 1px 3px rgba(0,0,0,0.1);">'
                    f'<div style="color:#6b7280;font-size:0.85em;font-weight:600;">{label}</div>'
                    f'<div style="color:{color};font-size:1.8em;font-weight:800;">{val}</div>'
                    f'</div>'
                ),
                unsafe_allow_html=True,
            )

    if not df_projects.empty and terlambat_proj > 0:
        st.markdown("<br/>", unsafe_allow_html=True)
        overdue_rows = df_projects[overdue_mask]
        overdue_items = "".join(
            f"<li><b>{r['Nama Project']}</b> — PIC: {r.get('PIC') or '-'}, "
            f"jatuh tempo {pd.to_datetime(r['Due Date']).strftime('%d/%m/%Y') if pd.notna(r['Due Date']) else '-'} "
            f"({render_project_status_badge(r['Status'])})</li>"
            for _, r in overdue_rows.iterrows()
        )
        st.markdown(
            (
                f'<div style="background:#fef2f2;border:1px solid #fecaca;border-radius:10px;'
                f'padding:14px 18px;margin-top:8px;">'
                f'<b style="color:#b91c1c;">⚠️ {terlambat_proj} Project Terlambat</b>'
                f'<ul style="margin:8px 0 0 0;">{overdue_items}</ul>'
                f'</div>'
            ),
            unsafe_allow_html=True,
        )

    if not df_projects.empty:
        st.markdown("<br/>", unsafe_allow_html=True)
        st.markdown("###### 📊 Progress per Project")
        st.caption("Progress bar berwarna sesuai Status (Belum Mulai = abu-abu, Berjalan = biru, Selesai = hijau, Tertunda = merah).")
        for _, r in df_projects.iterrows():
            if not str(r.get("Nama Project") or "").strip():
                continue
            st.markdown(render_project_progress_bar(r), unsafe_allow_html=True)

    st.markdown("<br/>", unsafe_allow_html=True)
    if not _GH_OK:
        st.warning(
            "⚠️ Backup GitHub tidak aktif, jadi Project Tracker HANYA tersimpan di server aplikasi ini dan "
            "bisa hilang kalau aplikasi restart/redeploy. Download backup di bawah ini setiap kali selesai "
            "update project, supaya bisa dipulihkan lagi kalau datanya hilang."
        )
    bcol1, bcol2 = st.columns([1, 2])
    with bcol1:
        st.download_button(
            "⬇️ Download Backup Project (Excel)",
            data=export_projects_excel(df_projects),
            file_name=f"backup_project_tracker_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with bcol2:
        up_projects_backup = st.file_uploader(
            "⬆️ Restore dari Backup (upload file hasil Download Backup di atas)",
            type=["xlsx", "xls"], key="up_projects_backup",
        )
        if up_projects_backup is not None:
            try:
                restored = import_projects_excel(up_projects_backup)
                _save_projects(restored)
                st.success(f"{len(restored)} project berhasil dipulihkan dari backup.")
                st.rerun()
            except Exception as e:
                st.error(f"Gagal membaca file backup: {e}")

    st.markdown("<br/>", unsafe_allow_html=True)
    st.markdown("###### Daftar Project")
    st.caption(
        "Tambah, edit, atau hapus baris langsung di tabel, lalu klik 💾 Simpan Perubahan. "
        "Mengetik/memilih di dalam tabel TIDAK akan me-refresh dashboard — perubahan baru diproses "
        "sekali saat tombol Simpan diklik, supaya lebih responsif. Isi Kendala/Catatan/Action Plan "
        "untuk mencatat hambatan dan rencana tindak lanjut per project."
    )

    with st.form("project_form", clear_on_submit=False):
        edited_df = st.data_editor(
            df_projects,
            num_rows="dynamic",
            use_container_width=True,
            key="projects_editor",
            column_config={
                "Nama Project": st.column_config.TextColumn("Nama Project", required=True, width="medium"),
                "Status": st.column_config.SelectboxColumn(
                    "Status", options=_PROJECT_STATUS_OPTIONS, required=True, width="small"
                ),
                "Due Date": st.column_config.DateColumn("Due Date", format="DD/MM/YYYY", width="small"),
                "PIC": st.column_config.TextColumn("PIC", width="small"),
                "Progress (%)": st.column_config.ProgressColumn(
                    "Progress (%)", min_value=0, max_value=100, format="%d%%", width="small"
                ),
                "Kendala": st.column_config.TextColumn("Kendala", width="medium"),
                "Catatan": st.column_config.TextColumn("Catatan", width="medium"),
                "Action Plan": st.column_config.TextColumn("Action Plan", width="medium"),
            },
        )
        submitted = st.form_submit_button("💾 Simpan Perubahan")

    if submitted:
        clean = edited_df.copy()
        clean = clean[clean["Nama Project"].notna() & (clean["Nama Project"].astype(str).str.strip() != "")]
        if "Status" in clean.columns:
            clean["Status"] = clean["Status"].fillna("Belum Mulai")
        if "Progress (%)" in clean.columns:
            clean["Progress (%)"] = pd.to_numeric(clean["Progress (%)"], errors="coerce").fillna(0).clip(0, 100)
        _save_projects(clean)
        st.success("Perubahan project berhasil disimpan.")
        st.rerun()

    if not df_projects.empty:
        st.markdown("<br/>", unsafe_allow_html=True)
        status_counts = df_projects["Status"].value_counts().reindex(_PROJECT_STATUS_OPTIONS).fillna(0)
        fig_proj = go.Figure(data=[go.Pie(
            labels=status_counts.index,
            values=status_counts.values,
            hole=0.5,
            marker=dict(colors=[_PROJECT_STATUS_COLORS.get(s, "#9ca3af") for s in status_counts.index]),
        )])
        fig_proj.update_layout(height=320, margin=dict(t=30, b=10, l=10, r=10), title="Distribusi Status Project")
        st.plotly_chart(fig_proj, use_container_width=True, key="chart_project_status")

st.markdown("---")
st.subheader("📦 Export Laporan Lengkap")
exp_col1, exp_col2 = st.columns(2)
with exp_col1:
    if st.button("📊 Buat Laporan PPTX", key="btn_gen_pptx"):
        with st.spinner("Membuat laporan PPTX..."):
            try:
                pptx_bytes = generate_pptx_report(
                    periode_label=periode_label,
                    quarter_period_label=quarter_period_label,
                    scoreboards=scoreboards,
                    pilar_summary=pilar_summary,
                    mc_summary=mc_summary,
                    df_ads=df_ads,
                    walkin_current=walkin_current,
                )
                st.session_state["_pptx_report"] = pptx_bytes
                st.success("Laporan PPTX berhasil dibuat.")
            except Exception as e:
                st.error(f"Gagal membuat laporan PPTX: {e}")
    if st.session_state.get("_pptx_report"):
        st.download_button(
            "⬇️ Unduh Laporan PPTX",
            data=st.session_state["_pptx_report"],
            file_name=f"Laporan_MFlash_{date.today().isoformat()}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key="dl_pptx",
        )
with exp_col2:
    if st.button("📄 Buat Laporan PDF", key="btn_gen_pdf"):
        with st.spinner("Membuat laporan PDF..."):
            try:
                pdf_bytes = generate_pdf_report(
                    periode_label=periode_label,
                    quarter_period_label=quarter_period_label,
                    scoreboards=scoreboards,
                    pilar_summary=pilar_summary,
                    mc_summary=mc_summary,
                    df_ads=df_ads,
                    walkin_current=walkin_current,
                )
                st.session_state["_pdf_report"] = pdf_bytes
                st.success("Laporan PDF berhasil dibuat.")
            except Exception as e:
                st.error(f"Gagal membuat laporan PDF: {e}")
    if st.session_state.get("_pdf_report"):
        st.download_button(
            "⬇️ Unduh Laporan PDF",
            data=st.session_state["_pdf_report"],
            file_name=f"Laporan_MFlash_{date.today().isoformat()}.pdf",
            mime="application/pdf",
            key="dl_pdf",
        )

st.markdown(
    """<div style="text-align:center;color:#9ca3af;font-size:0.8em;margin-top:24px;">
    Dashboard Omset MFlash — Internal Use Only
    </div>""",
    unsafe_allow_html=True,
)
