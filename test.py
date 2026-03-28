import requests

BOT_TOKEN = ""
url = f"https://api.telegram.org/bot{BOT_TOKEN}/getForumTopic"
params = {
    "chat_id": -1001234567890, # ID группы
    "message_thread_id": 123    # ID темы
}
response = requests.post(url, json=params).json()
topic_name = response["result"]["name"] if response.get("ok") else None
print(topic_name) # Например: "Заявки на материалы"
