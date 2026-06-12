from excel_processor import ExcelProcessor
import os

def test_4_xlsx():
    processor = ExcelProcessor()
    test_file = '4.xlsx'
    if not os.path.exists(test_file):
        print(f"File {test_file} not found.")
        return

    results = processor.process_file(test_file)
    for res in results:
        print(f"\n--- Sheet: {res['sheet_name']} ---")
        print(f"School: {res['school_info']['school_name']}, Grade: {res['school_info']['grade']}")
        print(f"Teacher: {res['teacher_name']}, Principal: {res['principal_name']}")
        print(f"Subjects found: {res['subjects']}")
        if "التربية الإسلامية" in res['subjects']:
            print("Found التربية الإسلامية!")
            first_student = res['students'][0]
            print(f"First student ({first_student['name']}):")
            for sub in res['subjects'][:3]:
                t1 = first_student.get(f"{sub}_ف1")
                t2 = first_student.get(f"{sub}_ف2")
                avg = first_student.get(f"{sub}_معدل")
                print(f"  {sub}: T1={t1}, T2={t2}, Avg={avg}")
        else:
            print("FAILED to find التربية الإسلامية")

if __name__ == "__main__":
    test_4_xlsx()
