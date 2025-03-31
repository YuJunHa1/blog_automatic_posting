from openai import OpenAI
import time
import os
from dotenv import load_dotenv
load_dotenv()

key = input("어떤 주제로 글을 작성하시겠습니까?")
reference = input("참고 자료를 넣어주세요.")

# GPT버전 선택
gpt_version = "gpt-4o-mini"
# 인증 키 입력
os.environ["OPENAI_API_KEY"] = os.getenv("OPENAI_API_KEY")

client = OpenAI()
query = reference + "의 내용을 참조해서 이 내용과 알맞는 2개의 20자 미만의 부제목을 뽑아줘."
sub_title = client.chat.completions.create(
  model=gpt_version,
  messages=[
    {"role": "developer", "content": "You are a helpful assistant."},
    {"role": "user", "content": query}
  ]
)
sub_title = sub_title.choices[0].message.content
sub_titles = sub_title.split("\n") 
time.sleep(2)

paragraph = []

for i in range(0, len(sub_titles)):
    query = key + "에 대해서 //" + reference + "//를 참조해서" + sub_titles[i] + "에 대한 내용을 400자 이내로 써줘."
    answer_con = client.chat.completions.create(model=gpt_version,messages=[
    {"role": "developer", "content": "You are a helpful assistant."},
    {"role": "user", "content": query}])
    answer_con = answer_con.choices[0].message.content
    paragraph.append(sub_titles[1] + "\n" + answer_con)
  
for k in range(0, len(paragraph)):
    print(paragraph[k])
