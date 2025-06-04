import pandas as pd
from openai import OpenAI # openai 라이브러리 v1.0.0 이상 기준

# --- 설정 (실제로는 안전하게 환경 변수 등으로 관리) ---
# client = OpenAI(api_key="YOUR_OPENAI_API_KEY")
# EXISTING_KNOWLEDGE_BASE = """
# 우리 학교의 교육 목표는 창의적 사고와 협업 능력 증진입니다.
# 주요 교육 활동은 프로젝트 기반 학습(PBL), 토론 수업, 코딩 교육입니다.
# 학생 평가는 과정 중심 평가를 지향하며, 수행평가와 포트폴리오를 활용합니다.
# 교사 연수는 주로 AI 활용 교육, 미래 교육 트렌드에 맞춰져 있습니다.
# """

def get_feedback_from_ai(teacher_input, knowledge_base):
    # 실제로는 API 키를 안전하게 관리해야 합니다.
    # 이 예제에서는 API 호출을 직접 실행하지 않고 개념만 보여드립니다.
    # client = OpenAI(api_key="YOUR_OPENAI_API_KEY") # 함수 내에서 client 초기화 또는 전역 client 사용

    prompt = f"""
    당신은 교육 데이터 분석 전문가 AI입니다.
    교사가 입력한 다음 내용을 분석해주세요: "{teacher_input}"

    참고로, 우리 학교의 기존 교육 자료 및 지침은 다음과 같은 내용을 주로 다룹니다:
    --- 기존 자료 ---
    {knowledge_base}
    --- 기존 자료 끝 ---

    교사가 입력한 내용이 기존 자료에서 다루지 않는 새롭거나 특이한 점이 있다면, 그 부분을 지적하고 어떤 검토나 수정이 필요할지 구체적으로 피드백을 제공해주세요.
    만약 기존 자료의 내용과 일관되거나 이미 다루고 있는 내용이라면 "기존 내용과 부합함" 또는 "특이사항 없음" 등으로 간략히 답변해주세요.
    피드백은 150자 이내로 요약해주세요.
    """
    print(f"--- AI에게 전달될 프롬프트 ---\n{prompt}\n--------------------------")

    # 실제 API 호출 부분 (주석 처리)
    # try:
    #     response = client.chat.completions.create(
    #         model="gpt-3.5-turbo", # 또는 gpt-4 등
    #         messages=[
    #             {"role": "system", "content": "당신은 교육 데이터 분석 전문가 AI입니다."},
    #             {"role": "user", "content": prompt}
    #         ],
    #         max_tokens=150,
    #         temperature=0.5
    #     )
    #     feedback = response.choices[0].message.content.strip()
    #     return feedback
    # except Exception as e:
    #     print(f"OpenAI API 호출 중 오류 발생: {e}")
    #     return "AI 피드백 생성 중 오류 발생"

    # 이 예제에서는 API 호출 대신 더미 응답을 반환합니다.
    if "블록체인" in teacher_input:
        return "블록체인 기술은 기존 교육 자료에서 명확히 다루지 않는 새로운 주제입니다. 교육적 활용 방안 및 실현 가능성에 대한 심층 검토가 필요합니다."
    elif "메타버스" in teacher_input and "코딩 교육" not in knowledge_base:
         return "메타버스 활용 교육은 혁신적이나, 기존 코딩 교육과의 연계성 및 인프라 구축 방안 검토가 필요합니다. 기존 자료에서는 깊이있게 다루지 않았습니다."
    else:
        return "기존 내용과 부합하거나 일반적인 내용으로 판단됩니다. 특이사항 없음."


def process_excel(filepath="teacher_data.xlsx"):
    try:
        df = pd.read_excel(filepath, header=None, sheet_name=0) # 첫번째 시트, 헤더 없음
    except FileNotFoundError:
        print(f"'{filepath}' 파일을 찾을 수 없습니다. 샘플 파일을 생성합니다.")
        # 샘플 데이터프레임 생성
        sample_data = {
            0: ["수업 아이디어: 양자컴퓨팅 기초", "학생 상담: A학생의 진로 고민"],
            1: ["연구 주제: 블록체인 기반 학습 이력 관리 시스템", "교사 연수 제안: 메타버스 활용 교육 워크숍"]
        }
        # 전치하여 1행에 교사 데이터, 2행은 비워둠
        df_sample_transposed = pd.DataFrame(sample_data).T
        # 2행을 비우기 위해 새로운 빈 행 추가 (실제로는 피드백으로 채워질 공간)
        empty_row = pd.Series([None] * len(df_sample_transposed.columns), name=len(df_sample_transposed))
        df = pd.concat([df_sample_transposed.iloc[[0]], pd.DataFrame([empty_row.values], columns=df_sample_transposed.columns, index=[1])])

        df.to_excel(filepath, index=False, header=False)
        print(f"샘플 파일 '{filepath}'가 생성되었습니다. 내용을 입력하고 다시 실행해주세요.")
        return

    if df.shape[0] < 1:
        print("파일에 데이터가 없습니다.")
        return

    # "기존 자료" 정의 (실제로는 외부 파일/DB에서 로드하거나, 더 정교하게 구성)
    EXISTING_KNOWLEDGE_BASE = """
    우리 학교의 교육 목표는 창의적 사고와 협업 능력 증진입니다.
    주요 교육 활동은 프로젝트 기반 학습(PBL), 토론 수업, 코딩 교육입니다.
    학생 평가는 과정 중심 평가를 지향하며, 수행평가와 포트폴리오를 활용합니다.
    교사 연수는 주로 AI 활용 교육, 미래 교육 트렌드에 맞춰져 있습니다.
    """

    # 1행 데이터 가져오기 (2열까지만 예시로)
    row1_data = df.iloc[0, :2].fillna("").astype(str).tolist() # 2열까지만, 없으면 빈 문자열

    # 2행을 피드백으로 채우기 (없으면 생성)
    if df.shape[0] < 2:
        df.loc[1] = [None] * df.shape[1]

    feedback_row = []
    for i, cell_data in enumerate(row1_data):
        if pd.isna(cell_data) or cell_data.strip() == "":
            feedback_row.append("") # 입력이 없으면 피드백도 없음
            if i < df.shape[1]: df.iloc[1, i] = ""
        else:
            print(f"\n[셀 {chr(65+i)}1 분석 중] 입력 내용: {cell_data}")
            feedback = get_feedback_from_ai(cell_data, EXISTING_KNOWLEDGE_BASE)
            print(f"[AI 피드백]: {feedback}")
            feedback_row.append(feedback)
            if i < df.shape[1]: df.iloc[1, i] = feedback # 2행에 피드백 작성
        
    # 나머지 열에 대해서도 빈칸 또는 기존 값 유지 (예시에서는 2열까지만 처리)
    for j in range(len(row1_data), df.shape[1]):
        if df.shape[0] > 1: # 2행이 존재하면
             df.iloc[1, j] = df.iloc[1, j] if pd.notna(df.iloc[1, j]) else "" # 기존 값 유지 또는 빈칸
        else: # 2행이 존재하지 않으면 빈칸 추가
             df.loc[1,j] = ""


    # 결과 저장
    output_filepath = filepath.replace(".xlsx", "_feedback.xlsx")
    df.to_excel(output_filepath, index=False, header=False)
    print(f"\n피드백이 추가된 파일이 '{output_filepath}'로 저장되었습니다.")

# --- 실행 ---
if __name__ == "__main__":
    # 실제 API 키 설정 (환경 변수 등 사용 권장)
    # import os
    # openai_api_key = os.getenv("OPENAI_API_KEY")
    # if not openai_api_key:
    #     print("OpenAI API 키가 설정되지 않았습니다. OPENAI_API_KEY 환경 변수를 설정해주세요.")
    # else:
    #     client = OpenAI(api_key=openai_api_key)
    #     process_excel()

    # API 키 없이 데모 실행을 위해 위 OpenAI client 초기화 부분 주석 처리하고 실행
    process_excel()