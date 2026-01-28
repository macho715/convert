# GWM-P13430-RP-001_Rev0  Loadout Analysis Report-3

**Source:** `GWM-P13430-RP-001_Rev0  Loadout Analysis Report-3.pdf`  
**Type:** PDF  
**Pages:** 17  
**Parsed At:** 2026-01-14T06:04:25.471095Z  

---

## Page 1

Great Waters Maritime L.L.C
Project:
SAMSUNG SPMT LOADOUT CALCULATION
Client:
Document Title:
LOADOUT ANALYSIS OF “LCT BUSHRA”
Document No: GWM-P13430-RP-001
0 13 Jan 2026 Issued for Review AR MM VS
Prepared Reviewed Approved Client
Rev Date Description
by by by Approval



### Table 1-0

| Great Waters Maritime L.L.C |  |  |  |  |  |  |
| --- | --- | --- | --- | --- | --- | --- |
| Project:
SAMSUNG SPMT LOADOUT CALCULATION |  |  |  |  |  |  |
| Client: |  |  |  |  |  |  |
| Document Title:
LOADOUT ANALYSIS OF “LCT BUSHRA”
Document No: GWM-P13430-RP-001 |  |  |  |  |  |  |
|  |  |  |  |  |  |  |
|  |  |  |  |  |  |  |
| 0 | 13 Jan 2026 | Issued for Review | AR | MM | VS |  |
| Rev | Date | Description | Prepared
by | Reviewed
by | Approved
by | Client
Approval |


## Page 2

Loadout Analysis of “LCT Bushra”
Doc No: GWM-P13430-RP-001 / Rev 0
13 Jan 2026
Revision Record Sheet
Revision Issue Date Purpose Description of Update
A 09-01-2026 Issued for Internal Review -
0 13-01-2026 Issued for Review -
Disclaimer
Great Waters Maritime LLC, executes all the projects using the best engineering expertise and best logical
judgments in performing analysis, design and other engineering services as requested by its clients. The drawings,
reports, findings, conclusions or any statements that Great Waters Maritime LLC, delivered are solely for the use of
its Client for whom it is contracted and are not intended for any other use including legal proceedings. It does its
utmost to support the needs of each of its Clients and presents its findings in an unbiased, objective manner using
its engineering resources. In the event that Great Waters Maritime LLC and its employees, either permanent,
temporary or contract, are brought into any court of law for any reasons for any legal proceedings whether is at fault
or not, Client acknowledges by using services of Great Waters Maritime LLC that Great Waters Maritime LLC shall
not be liable for any consequential damages and that Great Waters Maritime LLC is entitled to reimbursement by its
Client for all direct and indirect hours and expenses expended for such legal proceedings.
Page 2



### Table 2-0

|  | Revision |  |  | Issue Date |  |  | Purpose |  |  | Description of Update |  |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| A |  |  | 09-01-2026 |  |  | Issued for Internal Review |  |  | - |  |  |
| 0 |  |  | 13-01-2026 |  |  | Issued for Review |  |  | - |  |  |
|  |  |  |  |  |  |  |  |  |  |  |  |
|  |  |  |  |  |  |  |  |  |  |  |  |
|  |  |  |  |  |  |  |  |  |  |  |  |
|  |  |  |  |  |  |  |  |  |  |  |  |


## Page 3

Loadout Analysis of “LCT Bushra”
Doc No: GWM-P13430-RP-001 / Rev 0
13 Jan 2026
TABLE OF CONTENTS
1. INTRODUCTION ................................................................................................................... 4
2. SCOPE .................................................................................................................................. 4
3. GENERAL DATA ................................................................................................................... 4
3.1 Units ............................................................................................................................. 4
3.2 Constants ..................................................................................................................... 4
4. REFERENCE DOCUMENTS ................................................................................................. 5
4.1 Drawings and Documents ............................................................................................ 5
4.2 Codes & Standards ...................................................................................................... 5
5. SOFTWARE DESCRIPTION ................................................................................................. 6
5.1 Moses ........................................................................................................................... 6
6. VESSEL DETAILS AND LOADING CONDITION .................................................................. 7
6.1 Vessel Details ............................................................................................................... 7
6.2 Initial Loading Condition ............................................................................................... 8
6.3 Vessel Floating Condition ............................................................................................. 8
7. BALLAST SEQUENCE .......................................................................................................... 9
7.1 Ballast Sequence for Transformer-1 ............................................................................ 9
7.2 Ballast sequence for Transformer-2 ........................................................................... 10
8. Stability Assessment ............................................................................................................ 11
8.1 Intact Stability Criteria ................................................................................................ 11
8.2 Intact Stability Assessment Results ........................................................................... 13
9. LONGITUDINAL STRENGTH VERIFICATION ................................................................... 14
10. CONCLUSION ..................................................................................................................... 16
APPENDIX A - LOADOUT SEQUENCE FOR TRANSFORMER 2 ............................................ 17
APPENDIX B – MOSES COORDINATE SYSTEM ..................................................................... 21
APPENDIX C – SECTION MODULUS CALCULATION ............................................................. 24
APPENDIX D – MOSES MODEL PLOTS ................................................................................... 27
APPENDIX E –BARGE HYDROSTATICS .................................................................................. 32
APPENDIX F –MOSES OUTPUT ............................................................................................... 34
Page 3

## Page 4

Loadout Analysis of “LCT Bushra”
Doc No: GWM-P13430-RP-001 / Rev 0
13 Jan 2026
1. INTRODUCTION
Great Waters Maritime LLC has been contacted by Samsung to provide engineering
support for the load-out of two transformers, each weighing 240 MT, onto the vessel “LCT
Bushra”.
This report presents the engineering analyses associated with the load-out operation of the
two transformers.
.
2. SCOPE
The purpose of this document is to detail the loadout analysis for the two transformers on
“LCT Bushra”.
The following engineering analyses are covered in the report:
o Ballast requirement for the vessel during five stages of loadout operation.
o Verifying the longitudinal strength at each stage of the loadout operation.
o Verifying the intact stability criteria at each stage of loadout.
3. GENERAL DATA
3.1 Units
Loads or Forces Newton, kilo Newton [N, kN]
Mass kilogram [kg]
Length meter [m]
Area square meter [m²]
Section Modulus cubic meter [m3]
Moment of Inertia meter to the fourth [m4]
Moment Newton meter [Nm]
Kilo Newton meter [kNm]
Stress Newton/square meter [Pa, N/m²]
Newton/square millimeter [MPa, N/mm²]
3.2 Constants
Specific density of steel 7850 [kg/m3]
Gravity 9.81 [m/s²]
E-modulus of steel 2.06*1011 [N/m²]
Page 4

## Page 5

Loadout Analysis of “LCT Bushra”
Doc No: GWM-P13430-RP-001 / Rev 0
13 Jan 2026
4. REFERENCE DOCUMENTS
4.1 Drawings and Documents
• GWM-P13430-DA-001_Loadout Sequence Drawing_Rev.0
4.2 Codes & Standards
Table 4.2.1: Codes and Standards
No Document Number Document Title
1. AISC ASD Manual of Steel Construction – Allowable Stress
Design
2. GL Noble Denton 0030/ND Guidelines for Marine Transportation
3. DNV-ST-N001 Marine Operations and Marine Warranty Ed 2018
4. American Bureau of Rules for Building and Classing Steel Vessel, 1998-
Shipping ABS 1999
5. IMO A749 Resolution IMO Code on Intact Stability for All Types of Ships
Covered by IMO Instruments, 2002
Page 5



### Table 5-0

|  | No |  |  | Document Number |  |  | Document Title |  |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 1. |  |  | AISC ASD |  |  | Manual of Steel Construction – Allowable Stress
Design |  |  |
| 2. |  |  | GL Noble Denton 0030/ND |  |  | Guidelines for Marine Transportation |  |  |
| 3. |  |  | DNV-ST-N001 |  |  | Marine Operations and Marine Warranty Ed 2018 |  |  |
| 4. |  |  | American Bureau of
Shipping ABS |  |  | Rules for Building and Classing Steel Vessel, 1998-
1999 |  |  |
| 5. |  |  | IMO A749 Resolution |  |  | IMO Code on Intact Stability for All Types of Ships
Covered by IMO Instruments, 2002 |  |  |


## Page 6

Loadout Analysis of “LCT Bushra”
Doc No: GWM-P13430-RP-001 / Rev 0
13 Jan 2026
5. SOFTWARE DESCRIPTION
Following software(s) are utilized for barge stability assessment for tow.
5.1 Moses
MOSES is an acronym for Multi-Operational Structural Engineering Simulator. This is a
combined marine and structural package designed to describe a system and perform a
simulation including stress analysis at different phases of the simulation arising from marine
situation. MOSES accepts a 3-dimensional model of a vessel and/or a tubular structure. The
user can then combine the two structures in various ways and perform either static,
frequency domain or time domain analyses. The program computes hydrostatics, stability,
sea keeping, mooring, launching, and upending by using the same basic input.
The vessel model was developed based on the available structural drawings and general
arrangement.of the vessel Lightship particulars adopted in the analysis are in accordance
with the approved Stability Booklet. Wind and current projected areas were automatically
generated by MOSES.
Page 6

## Page 7

Loadout Analysis of “LCT Bushra”
Doc No: GWM-P13430-RP-001 / Rev 0
13 Jan 2026
6. VESSEL DETAILS AND LOADING CONDITION
6.1 Vessel Details
Name : LCT Bushra
Length Overall : 64.000 m
Length Between Perpendicular : 60.302 m
Breadth : 14.000 m
Depth : 3.650 m
Lightship Weight : 770.160 MT
LCG from AP : 26.350 m
VCG from Baseline : 3.880 m
TCG from centerline : -0.005 m (+ve to stbd)
Fig 6.1: Moses Model of LCT Bushra
Page 7

## Page 8

Loadout Analysis of “LCT Bushra”
Doc No: GWM-P13430-RP-001 / Rev 0
13 Jan 2026
6.2 Initial Loading Condition
The loading condition prior to the initiation of the load-out operation is summarized below.
Table 6.1: Initial Loading Condition
Weight LCG(1) TCG(1) VCG(1)
S.No. Description
(MT) (M) (M) (M)
1 LIGHTSHIP 770.16 26.350 -0.005 3.880
2 CREW EFFECTS 1.20 5.500 0.000 8.170
3 DO.P (43.7%) 1.26 11.243 -6.250 2.360
4 DO.S (43.7%) 1.26 11.243 6.250 2.360
5 FODB1.C (27.3%) 5.56 12.983 0.010 0.310
6 FODB1.P (43.5%) 5.64 12.693 -4.160 0.480
7 FODB1.S (68.4%) 8.87 12.533 4.220 0.610
8 FOW1.P (43.8%) 4.30 12.853 -6.250 1.620
9 FOW1.S (43.8%) 4.30 12.853 6.250 1.620
10 FWB2.P (100%) 110.66 50.093 -4.330 2.050
11 FWB2.S (100%) 110.66 50.093 4.330 2.050
12 FWCG1.P (10%) 14.85 42.543 -3.920 0.300
13 FWCG1.S (10%) 14.85 42.543 3.940 0.300
14 FWCG2.P(10%) 14.85 35.043 -3.920 0.300
15 FWCG2.S (10%) 14.85 35.043 3.940 0.300
16 LFRO.P (43.6%) 63.71 19.433 -3.930 0.900
17 LFRO.S (43.6%) 63.71 19.433 3.930 0.900
18 FW1.P (100%) 23.18 5.973 -6.080 3.100
19 FW1.S (100%) 23.18 5.973 6.080 3.100
20 FW2.P (100%) 13.89 0.113 -4.630 3.510
21 FW2.S (100%) 13.89 0.113 4.630 3.510
Total 1284.85 28.323 0.000 3.000
6.3 Vessel Floating Condition
Table 6.2: Vessel Floating Condition
Description Initial Loading Condition
Draft Aft (m) 2.46
Draft Fwd (m) 1.63
Draft Mid (m) 2.05
Trim (m) 0.83 (+ve by stern)
Heel (deg) 0.00
Page 8



### Table 8-0

|  |  |  | Description |  | Weight |  |  | LCG(1) |  |  | TCG(1) |  |  | VCG(1) |  |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
|  | S.No. |  |  | (MT) |  |  | (M) |  |  | (M) |  |  | (M) |  |  |
|  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| 1 |  |  | LIGHTSHIP | 770.16 |  |  | 26.350 |  |  | -0.005 |  |  | 3.880 |  |  |
| 2 |  |  | CREW EFFECTS | 1.20 |  |  | 5.500 |  |  | 0.000 |  |  | 8.170 |  |  |
| 3 |  |  | DO.P (43.7%) | 1.26 |  |  | 11.243 |  |  | -6.250 |  |  | 2.360 |  |  |
| 4 |  |  | DO.S (43.7%) | 1.26 |  |  | 11.243 |  |  | 6.250 |  |  | 2.360 |  |  |
| 5 |  |  | FODB1.C (27.3%) | 5.56 |  |  | 12.983 |  |  | 0.010 |  |  | 0.310 |  |  |
| 6 |  |  | FODB1.P (43.5%) | 5.64 |  |  | 12.693 |  |  | -4.160 |  |  | 0.480 |  |  |
| 7 |  |  | FODB1.S (68.4%) | 8.87 |  |  | 12.533 |  |  | 4.220 |  |  | 0.610 |  |  |
| 8 |  |  | FOW1.P (43.8%) | 4.30 |  |  | 12.853 |  |  | -6.250 |  |  | 1.620 |  |  |
| 9 |  |  | FOW1.S (43.8%) | 4.30 |  |  | 12.853 |  |  | 6.250 |  |  | 1.620 |  |  |
| 10 |  |  | FWB2.P (100%) | 110.66 |  |  | 50.093 |  |  | -4.330 |  |  | 2.050 |  |  |
| 11 |  |  | FWB2.S (100%) | 110.66 |  |  | 50.093 |  |  | 4.330 |  |  | 2.050 |  |  |
| 12 |  |  | FWCG1.P (10%) | 14.85 |  |  | 42.543 |  |  | -3.920 |  |  | 0.300 |  |  |
| 13 |  |  | FWCG1.S (10%) | 14.85 |  |  | 42.543 |  |  | 3.940 |  |  | 0.300 |  |  |
| 14 |  |  | FWCG2.P(10%) | 14.85 |  |  | 35.043 |  |  | -3.920 |  |  | 0.300 |  |  |
| 15 |  |  | FWCG2.S (10%) | 14.85 |  |  | 35.043 |  |  | 3.940 |  |  | 0.300 |  |  |
| 16 |  |  | LFRO.P (43.6%) | 63.71 |  |  | 19.433 |  |  | -3.930 |  |  | 0.900 |  |  |
| 17 |  |  | LFRO.S (43.6%) | 63.71 |  |  | 19.433 |  |  | 3.930 |  |  | 0.900 |  |  |
| 18 |  |  | FW1.P (100%) | 23.18 |  |  | 5.973 |  |  | -6.080 |  |  | 3.100 |  |  |
| 19 |  |  | FW1.S (100%) | 23.18 |  |  | 5.973 |  |  | 6.080 |  |  | 3.100 |  |  |
| 20 |  |  | FW2.P (100%) | 13.89 |  |  | 0.113 |  |  | -4.630 |  |  | 3.510 |  |  |
| 21 |  |  | FW2.S (100%) | 13.89 |  |  | 0.113 |  |  | 4.630 |  |  | 3.510 |  |  |
|  |  |  | Total | 1284.85 |  |  | 28.323 |  |  | 0.000 |  |  | 3.000 |  |  |




### Table 8-1

| Description | Initial Loading Condition |
| --- | --- |
| Draft Aft (m) | 2.46 |
| Draft Fwd (m) | 1.63 |
| Draft Mid (m) | 2.05 |
| Trim (m) | 0.83 (+ve by stern) |
| Heel (deg) | 0.00 |


## Page 9

Loadout Analysis of “LCT Bushra”
Doc No: GWM-P13430-RP-001 / Rev 0
13 Jan 2026
7. BALLAST SEQUENCE
The Ballast sequence for the loadout operation of Transformer-1 and Transformer-2 are
provided below. A minimum tide level of 1.8 m shall be maintained throughout the
entire load-out operation. Refer Appendix A for the Loadout sequence.
7.1 Ballast Sequence for Transformer-1
Table 7.1: Ballast Arrangement at each stage
FW1.P FW1.S FWB2.P FWB2.S FW2.P FW2.S
Stages Weight (MT)
% Filled % Filled % Filled % Filled % Filled % Filled
1 0 100 100 100 100 100 100
2 60 100 100 100 100 100 100
3 120 100 100 50 50 100 100
4 180 100 100 0 0 100 100
5 240 100 100 0 0 100 100
6 285 100 100 0 0 100 100
Notes:
1. Refer to Appendix-A for Loadout Sequence.
Table 7.2: Vessel Floating Condition at each stage
Draft Mid Trim Heel
Stages Weight (MT) Draft Aft (m) Draft Fwd (m)
(m) (m) (deg)
1 0 2.46 1.63 2.05 0.83 -0.07
2 60 2.30 1.92 2.11 0.38 -0.07
3 120 2.30 1.80 2.05 0.50 -0.07
4 180 2.30 1.68 1.99 0.62 -0.07
5 240 2.11 1.98 2.05 0.13 -0.07
6 285 1.97 2.20 2.09 -0.23 -0.08
Notes:
1. Refer to Appendix-F for Moses Results
Page 9



### Table 9-0

| Stages | Weight (MT) |  | FW1.P |  |  | FW1.S |  |  | FWB2.P |  |  | FWB2.S |  |  | FW2.P |  |  | FW2.S |  |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
|  |  |  | % Filled |  |  | % Filled |  |  | % Filled |  |  | % Filled |  |  | % Filled |  |  | % Filled |  |
| 1 | 0 | 100 |  |  | 100 |  |  | 100 |  |  | 100 |  |  | 100 |  |  | 100 |  |  |
| 2 | 60 | 100 |  |  | 100 |  |  | 100 |  |  | 100 |  |  | 100 |  |  | 100 |  |  |
| 3 | 120 | 100 |  |  | 100 |  |  | 50 |  |  | 50 |  |  | 100 |  |  | 100 |  |  |
| 4 | 180 | 100 |  |  | 100 |  |  | 0 |  |  | 0 |  |  | 100 |  |  | 100 |  |  |
| 5 | 240 | 100 |  |  | 100 |  |  | 0 |  |  | 0 |  |  | 100 |  |  | 100 |  |  |
| 6 | 285 | 100 |  |  | 100 |  |  | 0 |  |  | 0 |  |  | 100 |  |  | 100 |  |  |




### Table 9-1

| Stages | Weight (MT) | Draft Aft (m) | Draft Fwd (m) | Draft Mid
(m) | Trim
(m) | Heel
(deg) |
| --- | --- | --- | --- | --- | --- | --- |
| 1 | 0 | 2.46 | 1.63 | 2.05 | 0.83 | -0.07 |
| 2 | 60 | 2.30 | 1.92 | 2.11 | 0.38 | -0.07 |
| 3 | 120 | 2.30 | 1.80 | 2.05 | 0.50 | -0.07 |
| 4 | 180 | 2.30 | 1.68 | 1.99 | 0.62 | -0.07 |
| 5 | 240 | 2.11 | 1.98 | 2.05 | 0.13 | -0.07 |
| 6 | 285 | 1.97 | 2.20 | 2.09 | -0.23 | -0.08 |




### Table 9-2

| Draft Mid |
| --- |
| (m) |




### Table 9-3

| Trim |
| --- |
| (m) |




### Table 9-4

| Heel |
| --- |
| (deg) |


## Page 10

Loadout Analysis of “LCT Bushra”
Doc No: GWM-P13430-RP-001 / Rev 0
13 Jan 2026
7.2 Ballast sequence for Transformer-2
Table 7.3: Ballast Arrangement at each stage
FW1.P FW1.S FWB2.P FWB2.S FW2.P FW2.S
Stages Weight (MT)
% Filled % Filled % Filled % Filled % Filled % Filled
1 0 100 100 100 100 100 100
2 60 100 100 100 100 100 100
3 120 100 100 50 50 100 100
4 180 100 100 0 0 100 100
5 240 100 100 0 0 100 100
6 285 100 100 0 0 100 100
Notes:
1. Refer to Appendix-A for Loadout Plan.
Table 7.4: Vessel Floating Condition at each stage
Draft Mid Trim Heel
Stages Weight (MT) Draft Aft (m) Draft Fwd (m)
(m) (m) (deg)
1 0 2.69 1.98 2.34 0.71 -0.07
2 60 2.54 2.26 2.40 0.28 -0.07
3 120 2.54 2.14 2.34 0.40 -0.07
4 180 2.53 2.03 2.28 0.50 -0.08
5 240 2.38 2.31 2.35 0.07 -0.08
6 285 2.26 2.51 2.39 -0.25 -0.08
Notes:
1. Refer to Appendix-F for Moses Results.
Page 10



### Table 10-0

| Stages | Weight (MT) |  | FW1.P |  |  | FW1.S |  |  | FWB2.P |  |  | FWB2.S |  |  | FW2.P |  |  | FW2.S |  |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
|  |  |  | % Filled |  |  | % Filled |  |  | % Filled |  |  | % Filled |  |  | % Filled |  |  | % Filled |  |
| 1 | 0 | 100 |  |  | 100 |  |  | 100 |  |  | 100 |  |  | 100 |  |  | 100 |  |  |
| 2 | 60 | 100 |  |  | 100 |  |  | 100 |  |  | 100 |  |  | 100 |  |  | 100 |  |  |
| 3 | 120 | 100 |  |  | 100 |  |  | 50 |  |  | 50 |  |  | 100 |  |  | 100 |  |  |
| 4 | 180 | 100 |  |  | 100 |  |  | 0 |  |  | 0 |  |  | 100 |  |  | 100 |  |  |
| 5 | 240 | 100 |  |  | 100 |  |  | 0 |  |  | 0 |  |  | 100 |  |  | 100 |  |  |
| 6 | 285 | 100 |  |  | 100 |  |  | 0 |  |  | 0 |  |  | 100 |  |  | 100 |  |  |




### Table 10-1

| Stages | Weight (MT) | Draft Aft (m) | Draft Fwd (m) | Draft Mid
(m) | Trim
(m) | Heel
(deg) |
| --- | --- | --- | --- | --- | --- | --- |
| 1 | 0 | 2.69 | 1.98 | 2.34 | 0.71 | -0.07 |
| 2 | 60 | 2.54 | 2.26 | 2.40 | 0.28 | -0.07 |
| 3 | 120 | 2.54 | 2.14 | 2.34 | 0.40 | -0.07 |
| 4 | 180 | 2.53 | 2.03 | 2.28 | 0.50 | -0.08 |
| 5 | 240 | 2.38 | 2.31 | 2.35 | 0.07 | -0.08 |
| 6 | 285 | 2.26 | 2.51 | 2.39 | -0.25 | -0.08 |




### Table 10-2

| Draft Mid |
| --- |
| (m) |




### Table 10-3

| Trim |
| --- |
| (m) |




### Table 10-4

| Heel |
| --- |
| (deg) |


## Page 11

Loadout Analysis of “LCT Bushra”
Doc No: GWM-P13430-RP-001 / Rev 0
13 Jan 2026
8. Stability Assessment
8.1 Intact Stability Criteria
This section is intended to present the criteria regarding stability analysis of the barge.
A cargo barge shall have positive stability in calm water equilibrium position. In addition,
the system shall have sufficient dynamic stability (righting ability) to withstand the
overturning effect of the force produced by a steady wind from any horizontal direction.
In order to examine the stability of the system, criteria from GL Noble Denton 0030/ND,
Guidelines for Marine Transportation and IMO are considered as follows.
1) The critical metacentric height should not be less than 0.15 meter. (IMO Code on Intact
Stability for All Type of Ships Resolution A-749).
2) Range of Positive Stability Ø
The stability of the transportation barge shall be positive to a heel angle Ø beyond
equilibrium as given below:
• Heel angle, Ø will be greater than or equal to 40 degree, otherwise.
• Heel angle, Ø will be greater than or equal to the greater of 30 degree or
(R+15+15/GM) degree.
where,
R = Smaller or equal to the maximum dynamic heel angle in degree due to
wind and waves in design sea state for the tow route or to the heel
angle where the maximum righting moments occurs.
GM = Initial metacentric height, corrected for free surface effects in meters.
3) The area under righting lever curve up to the angle of maximum righting lever, or the
angle of down flooding, or 40 degrees, whichever is less, should not be less than 0.08
m-rad. (4.58 m-deg). (IMO Code on Intact Stability Special Criteria for Certain Types of
Ships -Resolution A.749 )
4) The area under the righting arm should not be less than 3.15 m-deg (0.055 metre-
radian) up to an angle of heel 30 degree. (IMO Code on Intact Stability For All Type Of
Ships Resolution A-749)
5) The area under the righting arm curve between the angles of heel of 30deg and 40deg,
or the down flooding angle, should not be less than 1.72 m-deg (0.03 metre-radian).
6) The righting lever should be at least 0.20 metres at a heel angle equal to or greater
than 30 degree. (IMO Code on Intact Stability for All Type of Ships Resolution A-749).
7) The maximum righting lever should occur at an angle of hell not less than 15 deg.
8) Dynamic Stability
The area under the righting moment curve will be greater than 1.4 times the area under
the wind heeling moment {(A+B)  1.4 (B+C)} as per the stability curve, Fig. 2.2
calculated up to an angle of heel corresponding to the second intercept of the two
curves.
The wind heeling moment is calculated on the basis of 1-minute mean wind speed. The
wind speed is 100 Knots.
Page 11

## Page 12

Loadout Analysis of “LCT Bushra”
Doc No: GWM-P13430-RP-001 / Rev 0
13 Jan 2026
Fig 8.1: Stability Curve for Intact Condition
Page 12

## Page 13

Loadout Analysis of “LCT Bushra”
Doc No: GWM-P13430-RP-001 / Rev 0
13 Jan 2026
8.2 Intact Stability Assessment Results
The results of the intact stability analysis are summarized below and the detailed
MOSES output is included in Appendix F.
Table 8.1: Intact Stability Results for Loadout of transformer-1
SL. Stage Stage Stage Stage Stage Stage
Particulars Required Remark
No. 1 2 3 4 5 6
1 Metacentric height (m) > 0.15 7.21 6.66 6.21 5.82 5.47 5.21 OK
Range of positive stability
2 > 40.0 >70 68.93 65.52 61.57 57.20 54.02 OK
(deg)
3 Wind area ratio > 1.40 >70 37.19 37.12 35.88 32.79 29.08 OK
Area at maximum righting
4 > 4.58 32.00 29.93 27.62 21.98 19.60 13.00 OK
arm (m-deg)
Heel angle at maximum
5 > 15.0 24.00 24.00 24.00 22.00 22.00 18.00 OK
righting arm (deg)
Area up to 30 deg heel angle
6 > 3.15 44.59 41.52 38.05 34.20 29.77 26.30 OK
(m-deg)
Area up to 40 deg heel angle
7 > 4.58 63.47 58.58 53.05 46.88 39.95 34.26 OK
(m-deg)
Area between 30 and 40 deg
8 > 1.72 18.88 17.06 15.00 12.68 10.18 7.96 OK
(m-deg)
The righting lever GZ at heel
9 > 0.20 2.1 1.80 1.70 1.60 1.40 1.20 OK
angle of 30 degree (m)
Table 8.2: Intact Stability Results for Loadout of transformer-2
SL. Stage Stage Stage Stage Stage Stage
Particulars Required Remark
No. 1 2 3 4 5 6
1 Metacentric height (m) > 0.15 5.77 5.45 5.09 4.77 4.48 4.29 OK
Range of positive stability
2 > 40.0 61.58 58.60 55.05 51.40 45.81 42.26 OK
(deg)
3 Wind area ratio > 1.40 49.02 48.50 46.33 41.82 36.88 30.61 OK
Area at maximum righting
4 > 4.58 21.77 20.26 18.49 12.11 8.80 7.68 OK
arm (m-deg)
Heel angle at maximum
5 > 15.0 22.00 22.00 22.00 18.00 16.00 16.00 OK
righting arm (deg)
Area up to 30 deg heel angle
6 > 3.15 34.09 31.36 28.17 24.50 20.30 16.53 OK
(m-deg)
Area up to 40 deg heel angle
7 > 4.58 46.79 42.44 37.33 31.52 25.02 19.12 OK
(m-deg)
Area between 30 and 40 deg
8 > 1.72 12.77 11.08 9.16 7.02 4.72 2.59 OK
(m-deg)
The righting lever GZ at heel
9 > 0.20 1.60 1.50 1.25 1.10 0.85 0.70 OK
angle of 30 degree (m)
Page 13



### Table 13-0

|  | SL. |  | Particulars | Required |  | Stage |  |  |  |  | Stage |  |  | Stage |  |  | Stage |  |  | Stage |  |  | Stage |  | Remark |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
|  | No. |  |  |  |  |  | 1 |  |  |  | 2 |  |  | 3 |  |  | 4 |  |  | 5 |  |  | 6 |  |  |
| 1 |  |  | Metacentric height (m) | > 0.15 | 7.21 |  |  |  |  | 6.66 |  |  | 6.21 |  |  | 5.82 |  |  | 5.47 |  |  | 5.21 |  |  | OK |
| 2 |  |  | Range of positive stability
(deg) | > 40.0 | >70 |  |  |  |  | 68.93 |  |  | 65.52 |  |  | 61.57 |  |  | 57.20 |  |  | 54.02 |  |  | OK |
| 3 |  |  | Wind area ratio | > 1.40 | >70 |  |  |  |  | 37.19 |  |  | 37.12 |  |  | 35.88 |  |  | 32.79 |  |  | 29.08 |  |  | OK |
| 4 |  |  | Area at maximum righting
arm (m-deg) | > 4.58 | 32.00 |  |  |  |  | 29.93 |  |  | 27.62 |  |  | 21.98 |  |  | 19.60 |  |  | 13.00 |  |  | OK |
| 5 |  |  | Heel angle at maximum
righting arm (deg) | > 15.0 | 24.00 |  |  |  |  | 24.00 |  |  | 24.00 |  |  | 22.00 |  |  | 22.00 |  |  | 18.00 |  |  | OK |
| 6 |  |  | Area up to 30 deg heel angle
(m-deg) | > 3.15 | 44.59 |  |  |  |  | 41.52 |  |  | 38.05 |  |  | 34.20 |  |  | 29.77 |  |  | 26.30 |  |  | OK |
| 7 |  |  | Area up to 40 deg heel angle
(m-deg) | > 4.58 | 63.47 |  |  |  |  | 58.58 |  |  | 53.05 |  |  | 46.88 |  |  | 39.95 |  |  | 34.26 |  |  | OK |
| 8 |  |  | Area between 30 and 40 deg
(m-deg) | > 1.72 | 18.88 |  |  |  |  | 17.06 |  |  | 15.00 |  |  | 12.68 |  |  | 10.18 |  |  | 7.96 |  |  | OK |
| 9 |  |  | The righting lever GZ at heel
angle of 30 degree (m) | > 0.20 | 2.1 |  |  |  |  | 1.80 |  |  | 1.70 |  |  | 1.60 |  |  | 1.40 |  |  | 1.20 |  |  | OK |




### Table 13-1

|  | SL. |  | Particulars | Required |  | Stage |  |  | Stage |  |  | Stage |  |  | Stage |  |  | Stage |  |  | Stage |  | Remark |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
|  | No. |  |  |  |  | 1 |  |  | 2 |  |  | 3 |  |  | 4 |  |  | 5 |  |  | 6 |  |  |
| 1 |  |  | Metacentric height (m) | > 0.15 | 5.77 |  |  | 5.45 |  |  | 5.09 |  |  | 4.77 |  |  | 4.48 |  |  | 4.29 |  |  | OK |
| 2 |  |  | Range of positive stability
(deg) | > 40.0 | 61.58 |  |  | 58.60 |  |  | 55.05 |  |  | 51.40 |  |  | 45.81 |  |  | 42.26 |  |  | OK |
| 3 |  |  | Wind area ratio | > 1.40 | 49.02 |  |  | 48.50 |  |  | 46.33 |  |  | 41.82 |  |  | 36.88 |  |  | 30.61 |  |  | OK |
| 4 |  |  | Area at maximum righting
arm (m-deg) | > 4.58 | 21.77 |  |  | 20.26 |  |  | 18.49 |  |  | 12.11 |  |  | 8.80 |  |  | 7.68 |  |  | OK |
| 5 |  |  | Heel angle at maximum
righting arm (deg) | > 15.0 | 22.00 |  |  | 22.00 |  |  | 22.00 |  |  | 18.00 |  |  | 16.00 |  |  | 16.00 |  |  | OK |
| 6 |  |  | Area up to 30 deg heel angle
(m-deg) | > 3.15 | 34.09 |  |  | 31.36 |  |  | 28.17 |  |  | 24.50 |  |  | 20.30 |  |  | 16.53 |  |  | OK |
| 7 |  |  | Area up to 40 deg heel angle
(m-deg) | > 4.58 | 46.79 |  |  | 42.44 |  |  | 37.33 |  |  | 31.52 |  |  | 25.02 |  |  | 19.12 |  |  | OK |
| 8 |  |  | Area between 30 and 40 deg
(m-deg) | > 1.72 | 12.77 |  |  | 11.08 |  |  | 9.16 |  |  | 7.02 |  |  | 4.72 |  |  | 2.59 |  |  | OK |
| 9 |  |  | The righting lever GZ at heel
angle of 30 degree (m) | > 0.20 | 1.60 |  |  | 1.50 |  |  | 1.25 |  |  | 1.10 |  |  | 0.85 |  |  | 0.70 |  |  | OK |


## Page 14

Loadout Analysis of “LCT Bushra”
Doc No: GWM-P13430-RP-001 / Rev 0
13 Jan 2026
9. LONGITUDINAL STRENGTH VERIFICATION
To check the longitudinal strength of the vessel, the shear force, bending moment and
deflection of the vessel are calculated and compared with the allowable shear stress,
bending stress and deflection.
It is necessary to have the shear area and section modulus of the vessel which is
calculated from the vessel construction drawings.
For all conditions, the allowable bending stress is assumed to be 175 MPa (as per ABS)
and for shear stress it is 110 MPa .
For shear area calculation, the watertight longitudinal bulkhead plates, side shell plates
of the vessel and longitudinal stiffeners are considered (based on construction drawings).
For vessel section modulus calculation, the deck and bottom plates are included on top
of that for shear area calculation. Addition to that side stiffeners and top and bottom
stiffeners are also included. The section modulus at the bottom is used, being the lesser
of the two. Refer detailed results and section modulus calculation in Appendix C
Table 9.1.: Allowable Bending moment and shear force for “LCT Bushra”
Shear Force Bending Moment
Allowable Shear Allowable Bending
11.0 kN/cm^2 17.5 kN/cm^2
Stress Stress
Shear Area 5686.14 cm2 Section Modulus 533 904.13 cm3
Allowable Shear Allowable Bending
6378.49 MT 9530.19 MT-m
Force Moment
Page 14



### Table 14-0

|  | Shear Force |  |  |  | Bending Moment |  |  |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Allowable Shear
Stress |  | 11.0 kN/cm^2 |  | Allowable Bending
Stress |  | 17.5 kN/cm^2 |  |
| Shear Area |  | 5686.14 cm2 |  | Section Modulus |  | 533 904.13 cm3 |  |
| Allowable Shear
Force |  | 6378.49 MT |  | Allowable Bending
Moment |  | 9530.19 MT-m |  |


## Page 15

Loadout Analysis of “LCT Bushra”
Doc No: GWM-P13430-RP-001 / Rev 0
13 Jan 2026
Table 9.2.: Bending Moment and Shear force for loadout of Transformer-1
Bending Moment (MT.m) Shear Force (MT)
Stages Weight (MT)
Maximum Allowable Maximum Allowable
1 0 4261 277
2 60 4819 296
3 120 5325 312
9530 6378
4 180 5790 326
5 240 6219 338
6 285 6519 346
Notes:
1. Refer to Appendix-F for Moses Results.
2. Refer to Appendix.C for Section Modulus Calculation
Table 9.3.: Bending Moment and Shear force for loadout of Transformer-2
Bending Moment (MT.m) Shear Force (MT)
Stages Weight (MT)
Maximum Allowable Maximum Allowable
1 0 3162 234
2 60 3517 251
3 120 3896 267
9530 6378
4 180 4446 281
5 240 4984 293
6 285 5366 302
Notes:
1. Refer to Appendix F for Moses Results.
2. Refer to Appendix.C for Section Modulus Calculation
Page 15



### Table 15-0

| Stages | Weight (MT) |  | Bending Moment (MT.m) |  |  |  |  |  | Shear Force (MT) |  |  |  |  |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
|  |  |  | Maximum |  |  | Allowable |  |  | Maximum |  |  | Allowable |  |
| 1 | 0 | 4261 |  |  | 9530 |  |  | 277 |  |  | 6378 |  |  |
| 2 | 60 | 4819 |  |  |  |  |  | 296 |  |  |  |  |  |
| 3 | 120 | 5325 |  |  |  |  |  | 312 |  |  |  |  |  |
| 4 | 180 | 5790 |  |  |  |  |  | 326 |  |  |  |  |  |
| 5 | 240 | 6219 |  |  |  |  |  | 338 |  |  |  |  |  |
| 6 | 285 | 6519 |  |  |  |  |  | 346 |  |  |  |  |  |




### Table 15-1

| Stages | Weight (MT) |  | Bending Moment (MT.m) |  |  |  |  |  | Shear Force (MT) |  |  |  |  |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
|  |  |  | Maximum |  |  | Allowable |  |  | Maximum |  |  | Allowable |  |
| 1 | 0 | 3162 |  |  | 9530 |  |  | 234 |  |  | 6378 |  |  |
| 2 | 60 | 3517 |  |  |  |  |  | 251 |  |  |  |  |  |
| 3 | 120 | 3896 |  |  |  |  |  | 267 |  |  |  |  |  |
| 4 | 180 | 4446 |  |  |  |  |  | 281 |  |  |  |  |  |
| 5 | 240 | 4984 |  |  |  |  |  | 293 |  |  |  |  |  |
| 6 | 285 | 5366 |  |  |  |  |  | 302 |  |  |  |  |  |


## Page 16

Loadout Analysis of “LCT Bushra”
Doc No: GWM-P13430-RP-001 / Rev 0
13 Jan 2026
10. CONCLUSION
The load-out analysis for loading of two transformers of 240 MT each onto the vessel “LCT
Bushra” is performed and the following conclusions are made.
• Intact Stability Analysis during the loadout of transformer-1 & transformer-2 has been
performed and is found to be satisfactory.
• Longitudinal strength assessment during the Loadout of transformer-1 &
transformer-2 has been conducted and is within allowable limits.
• A minimum tide of 1.8 m is required to commence the operation, and all
operations must be completed at or above this tide level.
Page 16

## Page 17

Loadout Analysis of “LCT Bushra”
Doc No: GWM-P13430-RP-001 / Rev 0
13 Jan 2026
APPENDIX A - LOADOUT SEQUENCE FOR TRANSFORMER 2
Page 17