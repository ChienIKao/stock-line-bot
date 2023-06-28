from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage
import os
from api.finance import Finance

TYPE = "1.ETF, 2.績優股, 3.金融股"

line_bot_api = LineBotApi(os.getenv("LINE_CHANNEL_ACCESS_TOKEN"))
line_handler = WebhookHandler(os.getenv("LINE_CHANNEL_SECRET"))
working_status = os.getenv("DEFALUT_TALKING", default = "true").lower() == "true"

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
    msg = event.message
    # global working_status
    if msg.type != "text":
        return
    
    if msg.text == "選股":
        reply = TYPE
    elif msg.text == "ETF" or msg.text == "etf" or msg.text == "1" or msg.text == "1.":
        reply = "ETF"
        for key, val in finance.ETF:
            reply += key + " : " + val + "\n"
        # reply = "00878.TW : 國泰永續高股息" + \
        #         "00919.TW : 群益台灣精選高息" + \
        #         "00713.TW : 元大台灣高息低波" + \
        #         "0056.TW : 元大高股息" + \
        #         "006208.TW : 富邦台灣50" + \
        #         "00690.TW : 兆豐藍籌30" + \
        #         "00850.TW : 元大臺灣ESG永續" + \
        #         "00701.TW : 國泰股利精選30"
    elif msg.text == "績優股" or msg.text == "2" or msg.text == "2.":
        reply = "績優股"
    elif msg.text == "金融股" or msg.text == "3" or msg.text == "3.":
        reply = "金融股"
    else:
        reply = "一袋米要扛幾樓"

    line_bot_api.reply_message(
        event.reply_token,
        TextSendMessage(text = reply)
    )
    return

    # if event.message.text in finance.ETF:
    #     symbol = finance.ETF
    # elif event.message.text in finance.BlueChip:
    #     symbol = finance.BlueChip
    # elif event.message.text in finance.Financial:
    #     symbol = finance.Financial

    # if event.message.text == "閉嘴":
    #     working_status = False
    #     line_bot_api.reply_message(
    #         event.reply_token,
    #         TextSendMessage(text="好的，我乖乖閉嘴 > <，如果想要我繼續說話，請跟我說 「說話」 > <"))
    #     return


if __name__ == "__main__":
    app.run()
