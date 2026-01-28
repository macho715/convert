# 날씨 데이터 수집 가이드라인

## 📋 개요

이 시스템은 API 대신 웹 검색을 통해 날씨 데이터를 수동으로 수집하여 사용하는 방식입니다.

## 🔍 검색할 웹사이트

### 1. UAE National Center of Meteorology (NCM)
- **URL**: https://www.ncm.ae
- **제공 정보**: 공식 날씨 예보, 샤말 경고, 해상 날씨
- **우선순위**: ⭐⭐⭐ (공식 데이터)

### 2. Windy.com
- **URL**: https://www.windy.com
- **좌표**: 24.12°N, 52.53°E (Mina Zayed Port)
- **제공 정보**: 풍속, 풍향, 파고, 가시거리
- **우선순위**: ⭐⭐⭐ (시각화 우수)

### 3. Meteoblue
- **URL**: https://www.meteoblue.com
- **제공 정보**: 상세 날씨 예보, 해상 조건
- **우선순위**: ⭐⭐

### 4. OpenWeatherMap
- **URL**: https://openweathermap.org
- **제공 정보**: 해상 날씨 데이터
- **우선순위**: ⭐⭐

## 📊 수집할 데이터 항목

| 항목 | 단위 | 설명 | 예시 | 필수 여부 |
|------|------|------|------|----------|
| Date | YYYY-MM-DD | 날짜 | 2026-01-18 | ✅ 필수 |
| Wind_Max_kn | knots | 최대 풍속 | 15.0 | ✅ 필수 |
| Gust_Max_kn | knots | 최대 돌풍 | 22.0 | ✅ 필수 |
| Wind_Dir_deg | degrees | 풍향 (0-360) | 315 (NW) | ✅ 필수 |
| Wave_Max_m | meters | 최대 파고 | 0.8 | ⚠️ 권장 |
| Visibility_km | km | 가시거리 | 8.0 | ⚠️ 권장 |
| Source | text | 데이터 출처 | "UAE NCM" | ⚠️ 권장 |
| Notes | text | 비고 | "Shamal detected" | 선택 |

## 🌪️ 샤말 바람 판단 기준

샤말(Shamal)은 UAE 지역의 특수한 기상 현상입니다:

- **방향**: NW (285-345도)
- **풍속**: ≥18kt 지속 또는 ≥22kt 돌풍
- **기간**: 보통 2-5일 지속
- **특징**: 먼지로 인한 가시거리 감소 (<6km)
- **예상 기간**: 2026년 2월 5일 ~ 14일

## 📝 데이터 입력 예시

```csv
Date,Wind_Max_kn,Gust_Max_kn,Wind_Dir_deg,Wave_Max_m,Visibility_km,Source,Notes
2026-01-18,15.0,20.0,315,0.6,10.0,Windy.com,Good weather
2026-02-05,20.0,28.0,300,1.2,4.0,UAE NCM,Shamal detected
2026-02-06,22.0,30.0,310,1.5,3.0,UAE NCM,Shamal continues
```

## 🔄 작업 순서

### 1단계: 템플릿 생성 (이미 완료)
```bash
python create_weather_data_template.py
```
→ `weather_data_template.csv` 파일 생성

### 2단계: 웹 검색 및 데이터 입력
1. 위 웹사이트에서 날씨 데이터 검색
2. Excel 또는 텍스트 에디터로 `weather_data_template.csv` 열기
3. 각 날짜별로 데이터 입력

### 3단계: JSON 변환
```bash
python convert_weather_csv_to_json.py
```
→ `weather_data_manual.json` 파일 생성

### 4단계: 히트맵 생성
```bash
python UntitSSSed-1.py
```
→ `AGI_TR_Weather_Risk_Heatmap_v2.png` 생성

## ⚙️ 설정 변경

`UntitSSSed-1.py` 파일에서 다음 설정을 변경할 수 있습니다:

```python
USE_MANUAL_JSON = True  # True: JSON 파일 사용, False: API 사용
WEATHER_JSON_PATH = "weather_data_manual.json"  # JSON 파일 경로
```

## 📌 주의사항

1. **데이터 정확도**: 가능한 한 공식 소스(UAE NCM) 우선 사용
2. **누락 데이터**: 일부 날짜 데이터가 없어도 보간 처리됨
3. **단위 확인**:
   - 풍속/돌풍: knots (kt)
   - 파고: meters (m)
   - 가시거리: kilometers (km)
   - 풍향: degrees (0-360, 북=0/360, 동=90, 남=180, 서=270)

## 🆘 문제 해결

### Q: JSON 파일을 찾을 수 없다는 오류
**A**: `convert_weather_csv_to_json.py`를 먼저 실행하세요.

### Q: 일부 날짜만 입력해도 되나요?
**A**: 네, 가능합니다. 누락된 날짜는 자동으로 보간됩니다.

### Q: API 모드로 다시 전환하려면?
**A**: `UntitSSSed-1.py`에서 `USE_MANUAL_JSON = False`로 설정하세요.

## 📞 참고 자료

- UAE NCM: https://www.ncm.ae
- Windy.com: https://www.windy.com
- 샤말 바람 정보: https://en.wikipedia.org/wiki/Shamal_(wind)

