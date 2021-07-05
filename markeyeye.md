0:종목코드(string
1:시간( ulong) - hhmm
2:대비부호(char)
'1' 상한
'2'상승
'3'보합
'4'하한
'5'하락
3:전일대비(long or float) - 주의) 반드시 대비부호(2)와 같이 요청을 하여야 함
4:현재가(long or float)
5:시가(long or float)
6:고가(long or float)
7:저가(long or float)
8:매도호가(long or float)
9:매수호가(long or float)
10:거래량( ulong)
11:거래대금(ulonglong) - 단위:원
12:장구분(char or empty)
'0' 장전
'1' 동시호가
'2' 장중
13:총매도호가잔량(ulong)
14:총매수호가잔량(ulong)
15:최우선매도호가잔량(ulong)
16:최우선매수호가잔량(ulong)
17:종목명(string)
20:총상장주식수(ulonglong) - 단위:주
21:외국인보유비율(float)
22:전일거래량(ulong)
23:전일종가(long or float)
24:체결강도(float)
25:체결구분(char or empty)
'1' 매수체결
'2' 매도체결
27:미결제약정(long)
28:예상체결가(long)
29:예상체결가대비(long) - 주의) 반드시 예샹체결가대비부호(30)와 같이 요청을 하여야 함
30:예상체결가대비부호(char or empty)
'1'상한
'2'상승
'3'보합
'4'하한
'5'하락
31:예상체결수량(ulong)
32:19일종가합(long or float)
33:상한가(long or float)
34:하한가(long or float)
35:매매수량단위(ushort)
36:시간외단일대비부호(char or empty)
'+'양수
'-'음수
37:시간외단일전일대비(long) - 주의) 반드시 시간외단일대비부호(36)와 같이 요청을 하여야 함
38:시간외단일현재가(long)
39:시간외단일시가(long)
40:시간외단일고가(long)
41:시간외단일저가(long)
42:시간외단일매도호가(long)
43:시간외단일매수호가(long)
44:시간외단일거래량(ulong)
45:시간외단일거래대금(ulonglong) - 단위:원
46:시간외단일총매도호가잔량(ulong)
47:시간외단일총매수호가잔량(ulong)
48:시간외단일최우선매도호가잔량(ulong)
49:시간외단일최우선매수호가잔량(ulong)
50:시간외단일체결강도(float)
51:시간외단일체결구분(char or empty)
'1'매수체결
'2'매도체결
53:시간외단일예상/실체결구분(char)
'1'예상체결
'2'실체결
54:시간외단일예상체결가(long)
55:시간외단일예상체결전일대비(long) - 주의) 반드시 시간외예상체결대비부호(56)와 같이 요청을 하여야 함
56:시간외단일예상체결대비부호(char or empty)
'1'상한
'2'상승
'3'보합
'4'하락
'5'하한
57:시간외단일예상체결수량(ulong)
59:시간외단일기준가(long)
60:시간외단일상한가(long)
61:시간외단일하한가(long)
62:외국인순매매(long)
63:52주최고가(long or float)
64:52주최저가(long or float)
65:연중주최저가(long or float)
66:연중최저가(long or float)
67:PER(float)
68:시간외매수잔량(ulong)
69:시간외매도잔량(ulong)
70:EPS(ulong)
71:자본금(ulonglong)- 단위:백만
72:액면가(ushort)
73:배당률(float)
74:배당수익률(float)
75:부채비율(float)
76:유보율(float)
77:자기자본이익률(float)
78:매출액증가율(float)
79:경상이익증가율(float)
80:순이익증가율(float)
81:투자심리(float)
82: VR(float)
83:5일 회전율(float)
84:4일 종가합(ulong)
85:9일 종가합(ulong)
86:매출액(ulonglong) - 단위: 백만
87:경상이익(ulonglong) - 단위:원
88:당기순이익(ulonglog) - 단위:원
89:BPS(ulong) - 주당순자산
90:영업이익증가율(float)
91:영업이익(ulonglong) - 단위:원
92:매출액영업이익률(float)
93:매출액경상이익률(float)
94:이자보상비율(float)
95:결산년월(ulong) - yyyymm
96:분기BPS(ulong) - 분기주당순자산
97:분기매출액증가율(float)
98:분기영업이액증가율(float)
99:분기경상이익증가율(float)
100:분기순이익증가율(float)
101:분기매출액(ulonglong) - 단위:원
102:분기영업이익(ulonglong) - 단위:원
103:분기경상이익(ulonglong) - 단위:원
104:분기당기순이익(ulonglong) - 단위:원
105:분개매출액영업이익률(float)
106:분기매출액경상이익률(float)
107:분기ROE(float) - 자기자본순이익률
108:분기이자보상비율(float)
109:분기유보율(float)
110:분기부채비율(float)
111:최근분기년월(ulong) - yyyymm
112:BASIS(float)
113:현지날짜(ulong) - yyyymmdd
114:국가명(string) - 해외지수 국가명
115:ELW이론가(ulong)
116:프로그램순매수(long)
117:당일외국인순매수잠정구분(char)
'\0'(0)해당없음
'1'확정
'2'잠정
118:당일외국인순매수(long)
119:당일기관순매수잠정구분(char)
'\0'(0)해당없음
'1'확정
'2'잠정
120:당일기관순매수(long)
121:전일외국인순매수(long)
122:전일기관순매수(long)
123:SPS(ulong)
124:CFPS(ulong)
125:EBITDA(ulong)
126:신용잔고율(float)
127:공매도수량(ulong)
128:공매도일자(ulong)
129:ELW e-기어링(float)
130:ELW LP보유양(ulong)
131:ELW LP보유율(float)
132:ELW LP Moneyness(float)
133:ELW LP Moneyness구분(char)
'1'ITM
'2'OTM
134:ELW 감마(float)
135:ELW 기어링(float)
136:ELW 내재변동성(float)
137:ELW 델타(float)
138:ELW 발행수량(ulong)
139:ELW 베가(float)
140:ELW 세타(float)
141:ELW 손익분기율(float)
142:ELW 역사적변동성(float)
143:ELW 자본지지점(float)
144:ELW 패리티(float)
145:ELW 프리미엄(float)
146:ELW 베리어(float)