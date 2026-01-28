Option A의 병렬 운영 패턴 분석 중:


[2 tools called]


**Option A의 병렬 운영 상세 분석**

## 병렬 운영 개념

SPMT 2개 세트가 동시에 서로 다른 작업을 수행해 전체 기간을 단축합니다.

---

## 1. 모빌라이제이션 단계 (병렬 준비)

```
타임라인:
2026-01-26: 1st set 모빌라이제이션 시작
            ↓
2026-01-26: 1st set 완료 (즉시 사용 가능)
            ↓
2026-02-04: 2nd set 모빌라이제이션 완료 (병렬 운영 시작)
```

- 1st set: 2026-01-26 완료 → 즉시 TR Units 1-2 작업 시작
- 2nd set: 2026-02-04 완료 → TR Units 3-4 작업 준비

---

## 2. 병렬 운영 시나리오 (시간대별)

### Phase 1: 1st set만 사용 (2026-01-29 ~ 2026-02-04)

| 날짜 | 1st set (SPMT) | 2nd set (SPMT) |
|------|----------------|----------------|
| 2026-01-29 | TR Unit 1 Load-out | 모빌라이제이션 중 |
| 2026-01-30 | TR Unit 2 Load-out | 모빌라이제이션 중 |
| 2026-02-01 | LCT 출항 (MZP → AGI) | 모빌라이제이션 중 |
| 2026-02-02 | AGI 입항 | 모빌라이제이션 중 |
| 2026-02-04 | TR Unit 1 Load-in (AGI) | 모빌라이제이션 완료 |

---

### Phase 2: 병렬 운영 시작 (2026-02-05 ~ 2026-02-13)

핵심: 두 세트가 동시에 다른 작업 수행

#### 2026-02-05 (병렬 운영 시작일)

| 시간대 | 1st set (SPMT) | 2nd set (SPMT) |
|--------|----------------|----------------|
| 오전 | TR Unit 1: AGI에서 TR Bay 4로 이동<br>Steel bridge 설치 | LCT가 MZP 도착<br>TR Units 3-4 준비 시작 |
| 오후 | TR Unit 1: SPMT에 적재<br>Transportation 시작 | TR Unit 3: MZP에서 SPMT 적재 준비 |

동시 작업:
- 1st set: AGI 현장에서 TR Unit 1 설치 작업
- 2nd set: MZP에서 TR Units 3-4 적재 준비

---

#### 2026-02-06 ~ 2026-02-07 (병렬 운영 정점)

| 날짜 | 1st set (AGI 현장) | 2nd set (MZP) |
|------|-------------------|---------------|
| 2026-02-06 | TR Unit 1: Turning 작업 (3일) | TR Unit 3: Load-out 준비<br>Beam Replacement |
| 2026-02-07 | TR Unit 1: Turning 계속 | TR Unit 3: Load-out 완료<br>TR Unit 4: 적재 시작 |

병렬 효과:
- AGI: 1st set로 TR Units 1-2 설치 진행
- MZP: 2nd set로 TR Units 3-4 적재 준비

---

#### 2026-02-08 ~ 2026-02-09 (병렬 운영 지속)

| 날짜 | 1st set (AGI) | 2nd set (MZP) |
|------|----------------|---------------|
| 2026-02-08 | TR Unit 1: Jack-down 완료<br>TR Unit 2: SPMT 적재 시작 | TR Unit 4: Load-out 완료<br>LCT 출항 준비 |
| 2026-02-09 | TR Unit 2: Transportation 시작 | LCT 출항 (MZP → AGI)<br>TR Units 3-4 운송 시작 |

---

#### 2026-02-10 ~ 2026-02-11 (병렬 운영 전환점)

| 날짜 | 1st set (AGI) | 2nd set (해상/AGI) |
|------|----------------|-------------------|
| 2026-02-10 | TR Unit 2: Turning 시작 | LCT 해상 운송 중 |
| 2026-02-11 | 1st set: SPMT shifting back to Port<br>(AGI → MZP 이동 시작) | LCT AGI 도착<br>TR Units 3-4 하역 준비 |

병렬 전환:
- 1st set: TR Units 1-2 작업 완료 → Port로 복귀
- 2nd set: TR Units 3-4를 AGI로 운송

---

### Phase 3: 2nd set 단독 운영 (2026-02-12 ~ 2026-02-22)

| 날짜 | 1st set | 2nd set (AGI) |
|------|---------|----------------|
| 2026-02-12 | Port 복귀 완료<br>대기 상태 | TR Unit 2: Jack-down 완료<br>TR Units 3-4: Load-in 시작 |
| 2026-02-13 | 대기 | TR Unit 3: Load-in 완료 |
| 2026-02-14 | 대기 | TR Unit 3: Steel bridge 설치<br>TR Unit 3: SPMT 적재 |
| 2026-02-18 | 대기 | TR Unit 4: SPMT 적재<br>TR Unit 4: Turning 시작 |
| 2026-02-20 | 대기 | 2nd set: SPMT shifting back to Port |
| 2026-02-22 | 대기 | TR Unit 4: Jack-down 완료 |

---

### Phase 4: 1st set 재활용 (2026-02-15 ~ 2026-03-03)

2nd set가 TR Units 3-4 작업 중, 1st set는 TR Units 5-6 작업 시작

| 날짜 | 1st set (MZP) | 2nd set (AGI) |
|------|---------------|---------------|
| 2026-02-15 | TR Units 5-6: 적재 준비 시작 | TR Unit 3: Turning 작업 중 |
| 2026-02-16 | TR Unit 5: Load-out | TR Unit 3: Turning 계속 |
| 2026-02-17 | TR Unit 6: Load-out | TR Unit 3: Jack-down 준비 |
| 2026-02-19 | LCT 출항 (MZP → AGI) | TR Unit 4: Turning 중 |
| 2026-02-20 | LCT 해상 운송 | 2nd set: Port 복귀 |
| 2026-02-21 | LCT AGI 도착 | (대기) |
| 2026-02-23 | TR Unit 5: SPMT 적재 | (대기) |
| 2026-02-27 | TR Unit 5: Jack-down 완료<br>TR Unit 6: SPMT 적재 | (대기) |
| 2026-03-01 | 1st set: Port 복귀 | (대기) |
| 2026-03-03 | (작업 완료) | (대기) |

---

### Phase 5: 2nd set 최종 재활용 (2026-02-24 ~ 2026-03-07)

1st set 작업 완료 후, 2nd set가 TR Unit 7 작업

| 날짜 | 1st set | 2nd set (MZP/AGI) |
|------|---------|-------------------|
| 2026-02-24 | (대기) | TR Unit 7: 적재 준비 |
| 2026-02-25 | (대기) | TR Unit 7: Load-out |
| 2026-02-27 | (대기) | LCT 출항 (MZP → AGI) |
| 2026-02-28 | (대기) | LCT AGI 도착 |
| 2026-03-01 | (대기) | TR Unit 7: Load-in |
| 2026-03-02 | (대기) | TR Unit 7: SPMT 적재 |
| 2026-03-04 | (대기) | TR Unit 7: Turning 시작 |
| 2026-03-06 | (대기) | 2nd set: Port 복귀 |
| 2026-03-07 | (대기) | TR Unit 7: Jack-down 완료 |

---

## 3. 병렬 운영의 핵심 메커니즘

### A. 시간 겹침 (Overlap)

```
2026-02-05 ~ 2026-02-11: 병렬 운영 기간 (7일)

1st set: AGI에서 TR Units 1-2 설치
         ↓
         [2026-02-05] ──────────── [2026-02-11] Port 복귀
         
2nd set: MZP에서 TR Units 3-4 적재
         ↓
         [2026-02-05] ──────────── [2026-02-11] AGI 도착
```

### B. 리소스 분리

| 리소스 | 1st set | 2nd set |
|--------|---------|---------|
| 위치 | AGI 현장 | MZP → AGI |
| 작업 | TR Units 1-2 설치 | TR Units 3-4 적재/운송 |
| 인력 | AGI 설치팀 | MZP 적재팀 |
| 장비 | SPMT 1st set | SPMT 2nd set |

### C. 순차적 전환

```
1st set 완료 → Port 복귀 → 다음 배치 준비
2nd set 완료 → Port 복귀 → 다음 배치 준비
```

---

## 4. 병렬 운영의 효과

### 시간 절약

| 항목 | 순차 운영 (예상) | 병렬 운영 (실제) | 절약 |
|------|------------------|------------------|------|
| TR Units 1-2 | 16일 | 16일 | - |
| TR Units 3-4 | 17일 (대기 포함) | 17일 (병렬) | 9일 |
| TR Units 5-6 | 17일 (대기 포함) | 17일 (병렬) | 9일 |
| TR Unit 7 | 12일 (대기 포함) | 12일 | - |
| 총 기간 | ~62일 | 40일 | 22일 절약 |

### 리소스 활용도

- SPMT 가동률: 순차 50% → 병렬 85%+
- 프로젝트 기간: 40일 (Option A) vs 62일 (Option B)

---

## 5. 병렬 운영의 제약사항

1. 초기 투자: SPMT 2세트 필요
2. 동기화: 두 세트 작업 일정 조율 필요
3. 리소스 분배: 인력/장비를 두 현장에 분산
4. 리스크: 한 세트 지연 시 전체 영향

---

## 요약

Option A의 병렬 운영은 SPMT 2개 세트를 동시에 사용해:
- 프로젝트 기간을 22일 단축 (40일 vs 62일)
- 리소스 활용도 향상
- 초기 투자 증가와 운영 복잡도 상승

이 방식은 속도 우선 전략입니다.