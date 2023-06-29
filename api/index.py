from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage, ImageSendMessage
from api.finance import Finance
# import os


# ETF, 績優股, 金融股
TYPE = "1. ETF \n2. 績優股 \n3. 金融股"
ETF = {
    '00878.TW': '國泰永續高股息', 
    '00919.TW': '群益台灣精選高息', 
    '00713.TW': '元大台灣高息低波', 
    '0056.TW': '元大高股息', 
    '006208.TW': '富邦台灣50', 
    '00690.TW': '兆豐藍籌30', 
    '00850.TW': '元大臺灣ESG永續', 
    '00701.TW': '國泰股利精選30'
}
BLUE_CHIP = {
    '2356.TW': '英業達',
    '2330.TW': '台積電', 
    '2454.TW': '聯發科', 
    '2357.TW': '華碩', 
    '6505.TW': '台塑',
    '1102.TW': '亞泥',
    '2324.TW': '仁寶'
}
FINANCIAL = {
    '2890.TW': '永豐金',
    '2886.TW': '兆豐金',
    '2885.TW': '元大金'
}

line_bot_api = LineBotApi('CnxTIV3ZENKBF4uLOFI2x2I2wwG7Y0ILmp0pR+TvHbE/pbTPpTxw3ea5qrfsfB/T4xnXZdwuBZHgFK+eXz/bE86B8Ge+YBtEt6mEduMjFf5Pi/VsNv5PrUkgK+AtTFKAKF1H05phg7v3dkKtDuSzYgdB04t89/1O/w1cDnyilFU=')
line_handler = WebhookHandler('11ce307d39f4e16e81dc9c49c3353ca9')
# line_bot_api = LineBotApi(os.getenv("LINE_CHANNEL_ACCESS_TOKEN"))
# line_handler = WebhookHandler(os.getenv("LINE_CHANNEL_SECRET"))
# working_status = os.getenv("DEFALUT_TALKING", default = "true").lower() == "true"

app = Flask(__name__)
finance = Finance()

# domain root
@app.route('/')
def home():
    return 'Hello, World!'

@app.route("/webhook", methods=['POST'])
def callback():
    # get X-Line-Signature header value
    signature = request.headers['X-Line-Signature']
    # get request body as text
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)
    # handle webhook body
    try:
        line_handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'


@line_handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    message = []
    msg = event.message.text
    if event.message.type != "text":
        return
    
    if msg == "選股":
        reply = TYPE
        message.append(TextSendMessage(text = reply))
    elif msg == "ETF" or msg == "etf" or msg == "1" or msg == "1.":
        reply = "ETF: \n"
        for key, val in ETF.items():
            reply += key + " : " + val + "\n"
        message.append(TextSendMessage(text = reply))
    elif msg == "績優股" or msg == "2" or msg == "2.":
        reply = "績優股: \n"
        for key, val in BLUE_CHIP.items():
            reply += key + " : " + val + "\n"
        message.append(TextSendMessage(text = reply))
    elif msg == "金融股" or msg == "3" or msg == "3.":
        reply = "金融股: \n"
        for key, val in FINANCIAL.items():
            reply += key + " : " + val + "\n"
        message.append(TextSendMessage(text = reply))
    elif msg in ETF or msg in BLUE_CHIP or msg in FINANCIAL:
        if msg in ETF:
            reply = finance.getReplyMsg(msg, ETF[msg])
            img_url = finance.getImg(msg)
        elif msg in BLUE_CHIP:
            reply = finance.getReplyMsg(msg, BLUE_CHIP[msg])
            img_url = finance.getImg(msg)
        elif msg in FINANCIAL:
            reply = finance.getReplyMsg(msg, FINANCIAL[msg])
            img_url = finance.getImg(msg)
        else:
            reply = '抱歉，請再試一次'

        message.append(TextSendMessage(text = reply))
        message.append(ImageSendMessage(
            original_content_url = img_url,
            preview_image_url = img_url
        ))
    else:
        reply = "我不知道你在說什麼"

    line_bot_api.reply_message(event.reply_token, message)

    return 

if __name__ == "__main__":
    app.run()
