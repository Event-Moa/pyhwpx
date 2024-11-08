
from pyhwpx import Hwp
from collections import Counter

hwp = Hwp()
hwp.Open(r"C:\Users\thdco\origin.hwp")

hwp.set_field_by_bracket()

datetime_list = {"4.8.(월)" : ["09:00","13:30","14:00"],
             "4.9.(화)" : ["10:00","14:00"],
             "4.10.(수)" : [" "],
             "4.11.(목)" : ["09:30","10:00","10:00","13:00","13:00","14:00","14:00"],
             "4.12.(금)" : ["09:00","09:00","10:00","11:00","11:00","14:00","15:30"],
             "4.13.(토)" : ["08:00","10:00","10:00","10:00","15:00","15:00"],
             "4.14.(일)" : [" "]}

col_num = 0
for value in datetime_list.values():
    for i in value:
      col_num += 1
      
date_list = []
for key in datetime_list.keys() : 
  for value in datetime_list[key] :
    date_list.append(key)
    
time_list = []
for value in datetime_list.values() : 
  for num in value:
    time_list.append(num)
    
event_list = ['중앙재난안전대책본부 영상회의(의료계 파업)','찾아가는 치매조기검진 및 치매 예방교육 ','·2024년 유성 데이터 기반 실증 리빙랩 발대식',
              '나도갈래 과학소풍','행복한 문화학교「박완서의 작품 세계 엿보기」','','아가맘 바라온 만남 「동화 태교와 태담 태교」',
              '상반기 공감 인문학 「임홍택 작가 초청 강연」','독서동아리 역량강화교육','제44회 장애인의 날 및 장애인종합복지관 개관 19주년 기념행사',
              '유성구 신속대응반 도상훈련 교육','청년기본계획수립 연구용역 업체 제안서 평가위원회','2024년⌜너·나·들·이⌟ 치매예방교실 1기',
              '2024년 공무직 산업시찰','중앙재난안전대책본부 영상회의(의료계 파업)','우리동네 도서관 프로젝트「이유리 작가 찾아가는 작가 강연」',
              '대한민국 창조경영 2024 시상식 참석','유성구 보훈회관 건립 관련 보훈단체 간담회','마음의 병과 치유 「아픈 마음의 연원과 철학 상담」',
              '노인맞춤 돌봄(응급안전안심) 지침교육','·2024 대청호 벚꽃길 마라톤대회','소셜다이닝「1인 가구 함께하는 식사」(1회차) ',
              '2024년 특성화 사업 제1회차‘인문학 콘서트’','행복한 문화학교 「햄스터 로봇 코딩」','노은도서관 동화 들려주기(Story Time)',
              '독서프로그램 「귀 기울여 영어동화」 영어원서 읽기','']

location_list = ['재난안전상황실','외삼1통경로당','대회의실','대전정보문화산업진흥원','유성도서관','','아가랑도서관','구암도서관','온라인(ZOOM)','유성구장애인종합복지관',
         '유성보건소', '중회의실','노인복지관','군산선유도','재난안전상황실','노은도서관','더플라자호텔(서울)','유성구커뮤니티센터(대회의실)','유성도서관','대회의실',
         '동구 신당동282', '하기동 하하팜농장', '유성문화원 1층전시실','유성도서관','노은도서관','어린이 영어마을 도서관','']

personnel_list = ['4명','16명','40명','20명','40명','','10명','50명','15명','800명','24명','7명','12명','85명','4명','60명','6명','21명','60명','30명','2500명','30명',
              '30명','14명','40명','10명','']

department_list = ['재난안전과','예방의약과','미래전략과','교육과학과','도서관운영과','','도서관운영과','도서관운영과','도서관운영과','사회돌봄과',
              '건강정책과','미래전략과','예방의약과','운영지원과','재난안전과','도서관운영과','홍보실','사회돌봄과','도서관운영과','사회돌봄과',
              '홍보실','미래전략과','문화관광과','도서관운영과','도서관운영과','도서관운영과','']

note_list = ['','','','','','','','','','','','','','','','','','','','','','','','','','','']

hwp.get_into_nth_table(1, select=True)  # 두 번째 표로 이동

hwp.TableLowerCell()
hwp.TableLowerCell()
hwp.TableCellBlockExtend()
hwp.TableColEnd()
hwp.TableColPageDown()


# 선택한 빈 행 잘라내기
hwp.Cut(remove_cell=True)
hwp.CloseEx()

hwp.MoveDocEnd()
for i in range(col_num):
    hwp.Paste()

hwp.put_field_text("date", date_list)
hwp.put_field_text("time", time_list)



hwp.put_field_text("event_name", event_list)
hwp.put_field_text("location", location_list)
hwp.put_field_text("personnel", personnel_list)
hwp.put_field_text("department", department_list)
hwp.put_field_text("note", note_list)

    
hwp.get_into_nth_table(1)

while hwp.TableMergeTable():
    pass


hwp.save_as(r"C:\Users\thdco\filled_template.hwp", "HWP", "forceopen:true")
