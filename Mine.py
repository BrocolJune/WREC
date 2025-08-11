import pandas as pd
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut, GeocoderUnavailable
import time

def get_location(address):
    """
    Nominatim을 사용하여 주소를 위경도 좌표로 변환
    """
    # 한국 주소임을 명시하기 위해 'South Korea' 추가
    search_address = f"{address}, South Korea"
    
    # Nominatim 초기화 (user_agent는 필수)
    geolocator = Nominatim(user_agent="my_geocoder")
    
    try:
        # 위치 검색
        location = geolocator.geocode(search_address)
        
        if location:
            return location.latitude, location.longitude
        else:
            print(f"주소를 찾을 수 없습니다: {address}")
            return None, None
            
    except (GeocoderTimedOut, GeocoderUnavailable) as e:
        print(f"Error with address {address}: {str(e)}")
        time.sleep(1)  # 오류 발생시 1초 대기
        return None, None
    except Exception as e:
        print(f"Unexpected error with address {address}: {str(e)}")
        return None, None

def main():
    # Excel 파일 읽기
    file_path = r"C:\Users\user\Downloads\Korea_Mine_Metal.xlsx"  # 실제 Excel 파일 경로로 수정하세요
    df = pd.read_excel(file_path)
    
    # 새로운 열 생성
    df["위도"] = None
    df["경도"] = None
    
    total = len(df)
    # 각 주소에 대해 위경도 변환
    for idx, row in df.iterrows():
        address = row["소재지"]
        print(f"Processing ({idx+1}/{total}): {address}")
        
        # 위경도 가져오기
        lat, lon = get_location(address)
        
        # 결과 저장
        df.at[idx, "위도"] = lat
        df.at[idx, "경도"] = lon
        
        # API 호출 제한을 위한 딜레이
        time.sleep(1)  # Nominatim은 초당 1회 요청 제한
        
        # 중간 저장 (100개마다)
        if (idx + 1) % 100 == 0:
            print(f"중간 저장 중... ({idx + 1}/{total})")
            temp_output = r"C:\Users\user\Downloads\Korea_Mine_Metal_geocoded_temp.xlsx"
            df.to_excel(temp_output, index=False)
        
        # 진행상황 출력 (100개마다)
        if (idx + 1) % 100 == 0:
            print(f"Processed {idx + 1} addresses...")
    
    # 결과를 Excel 파일로 저장
    output_path = r"C:\Users\user\Downloads\Korea_Mine_Metal_geocoded.xlsx"  # 결과 저장할 경로를 입력하세요
    df.to_excel(output_path, index=False)
    print(f"Results saved to {output_path}")

if __name__ == "__main__":
    main()
