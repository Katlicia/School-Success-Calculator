import pandas as pd
import numpy as np
import random

def create_program_outcomes():
    """Program çıktıları için örnek veri oluştur"""
    # Rastgele sayıda program çıktısı (4-8 arası)
    num_outcomes = random.randint(4, 8)
    
    base_outcomes = [
        "Matematik, fen bilimleri ve kendi dalları ile ilgili mühendislik konularında yeterli bilgi birikimi",
        "Problem tanımlama, modelleme ve çözme becerisi",
        "Bir sistemi, sistem bileşenini ya da süreci analiz etme becerisi",
        "Disiplin içi ve çok disiplinli takımlarda etkin biçimde çalışabilme becerisi",
        "Mühendislik problemlerini saptama, tanımlama, formüle etme ve çözme becerisi",
        "Mesleki ve etik sorumluluk bilinci",
        "Etkin iletişim kurma becerisi",
        "Mühendislik uygulamalarının evrensel ve toplumsal boyutlarda etkisini kavrama"
    ]
    
    # Rastgele seçilen sayıda çıktı al
    outcomes = random.sample(base_outcomes, num_outcomes)
    
    # DataFrame oluştur
    df = pd.DataFrame(outcomes)
    df.to_excel('program_ciktilari.xlsx', index=False, header=False)
    print(f"{num_outcomes} adet program çıktısı oluşturuldu: program_ciktilari.xlsx")
    return num_outcomes

def create_course_outcomes():
    """Ders çıktıları için örnek veri oluştur"""
    # Rastgele sayıda ders çıktısı (3-6 arası)
    num_outcomes = random.randint(3, 6)
    
    base_outcomes = [
        "Temel programlama kavramlarını anlama ve uygulama",
        "Algoritma geliştirme ve problem çözme becerisi",
        "Nesne yönelimli programlama prensiplerini uygulama",
        "Veri yapılarını anlama ve kullanma",
        "Temel veritabanı işlemlerini gerçekleştirme",
        "Web teknolojilerini kullanabilme becerisi"
    ]
    
    # Rastgele seçilen sayıda çıktı al
    outcomes = random.sample(base_outcomes, num_outcomes)
    
    # DataFrame oluştur
    df = pd.DataFrame(outcomes)
    df.to_excel('ders_ciktilari.xlsx', index=False, header=False)
    print(f"{num_outcomes} adet ders çıktısı oluşturuldu: ders_ciktilari.xlsx")
    return num_outcomes

def create_table1(num_program_outcomes, num_course_outcomes):
    """Tablo 1 için örnek veri oluştur (Program Çıktıları - Ders Çıktıları İlişkisi)"""
    data = []
    
    for i in range(num_program_outcomes):
        row = [f"PÇ{i+1}"]  # Program çıktısı numarası
        # Ders çıktıları için değerler (0-1 arası)
        values = np.round(np.random.uniform(0, 1, num_course_outcomes), 2)
        row.extend(values)
        data.append(row)
    
    columns = ['PÇ No'] + [f'ÇK{i+1}' for i in range(num_course_outcomes)]
    df = pd.DataFrame(data, columns=columns)
    df.to_excel('tablo1.xlsx', index=False)
    print("Tablo 1 oluşturuldu: tablo1.xlsx")

def create_table2(num_course_outcomes):
    """Tablo 2 için örnek veri oluştur (Değerlendirmeler - Ders Çıktıları İlişkisi)"""
    # Rastgele sayıda değerlendirme kriteri oluştur (3-5 arası)
    num_criteria = random.randint(3, 5)  # Maksimum 5 kriter
    
    # Temel kriterler
    criteria_groups = {
        "Ödev": ["Ödev", "Ödev 2"],
        "Quiz": ["Quiz"],
        "Proje": ["Proje"],
        "Vize": ["Vize"],
        "Final": ["Final"]
    }
    
    # Ders çıktısı içerikleri
    course_outcomes = [
        "Temel programlama kavramlarını anlama ve uygulama",
        "Algoritma geliştirme ve problem çözme becerisi",
        "Nesne yönelimli programlama prensiplerini uygulama",
        "Veri yapılarını anlama ve kullanma",
        "Temel veritabanı işlemlerini gerçekleştirme",
        "Web teknolojilerini kullanabilme becerisi"
    ]
    
    # Rastgele kriter grupları seç
    selected_groups = random.sample(list(criteria_groups.keys()), min(num_criteria, len(criteria_groups)))
    selected_groups.sort()  # Grupları sırala
    
    # Seçilen gruplardan kriterleri oluştur
    selected_criteria = []
    for group in selected_groups:
        # Her grup için kaç kriter alınacağını belirle (1 ile o gruptaki maksimum kriter sayısı arasında)
        num_group_criteria = random.randint(1, len(criteria_groups[group]))
        # Gruptan sıralı olarak kriter seç
        group_criteria = criteria_groups[group][:num_group_criteria]
        selected_criteria.extend(group_criteria)
        
        # Toplam kriter sayısı 5'i geçerse, fazla kriterleri kaldır
        if len(selected_criteria) > 5:
            selected_criteria = selected_criteria[:5]
            break
    
    # Ders çıktılarından rastgele seç
    selected_outcomes = random.sample(course_outcomes, num_course_outcomes)
    
    data = []
    for i in range(num_course_outcomes):
        row = [f"ÇK{i+1}", selected_outcomes[i]]  # Ders çıktısı numarası ve içeriği
        # Değerlendirme kriterleri için değerler (0-1 arası)
        values = np.round(np.random.uniform(0, 1, len(selected_criteria)), 2)
        row.extend(values)
        data.append(row)
    
    columns = ['ÇK No', 'Ders Çıktısı'] + selected_criteria
    df = pd.DataFrame(data, columns=columns)
    df.to_excel('tablo2.xlsx', index=False)
    print(f"Tablo 2 oluşturuldu ({len(selected_criteria)} kriter): tablo2.xlsx")
    return selected_criteria

def create_student_grades(criteria):
    """Öğrenci notları için örnek veri oluştur"""
    # Rastgele sayıda öğrenci (4-15 arası)
    num_students = random.randint(4, 15)
    
    # Öğrenci numaralarını oluştur (220502000'den başlayarak)
    base_number = 220502000
    # Rastgele numaralar üret (tekrar etmeyen)
    student_numbers = random.sample(range(base_number, base_number + 1000), num_students)
    student_numbers.sort()  # Numaraları sırala
    
    data = []
    for student_no in student_numbers:
        # Her değerlendirme için 0-100 arası rastgele notlar
        grades = np.random.randint(40, 101, len(criteria))
        row = [str(student_no)] + list(grades)
        data.append(row)
    
    # Tablo 2'deki kriterlerle aynı sırada olmasını sağla
    columns = ['Öğrenci'] + criteria
    df = pd.DataFrame(data, columns=columns)
    df.to_excel('ogrenci_notlari.xlsx', index=False)
    print(f"{num_students} öğrenci için notlar oluşturuldu: ogrenci_notlari.xlsx")

print("Test verileri oluşturuluyor...")
num_program_outcomes = create_program_outcomes()
num_course_outcomes = create_course_outcomes()
create_table1(num_program_outcomes, num_course_outcomes)
criteria = create_table2(num_course_outcomes)
create_student_grades(criteria)
print("\nTüm test verileri oluşturuldu!")
