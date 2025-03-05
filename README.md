# Sheet2SQL
Python-based project that reads an Excel file and updates specific target columns in a database


# 가상환경 생성
python -m venv excel-db-env

# 가상환경 활성화 (Windows)
excel-db-env\Scripts\activate

# 가상환경 활성화 (macOS/Linux)
source excel-db-env/bin/activate

# 필요한 패키지 설치
pip install pandas openpyxl tk mysql-connector-python

# 프로그램 실행
python main.py