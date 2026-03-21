from gastos_utils import *
from telegram.ext import ApplicationBuilder, MessageHandler, filters, CommandHandler
from telegram import Update
from telegram.ext import CallbackQueryHandler


def main():
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, manejar_mensaje))
    app.add_handler(CallbackQueryHandler(manejar_boton))
    app.add_handler(CommandHandler("cancelar", cancelar)) 
    app.add_handler(CommandHandler("saldo", saldo)) 
    app.add_handler(CommandHandler("pagos", pagos)) 
    app.add_handler(CommandHandler("dolar", dolar)) 
    app.add_handler(CommandHandler("ultimos", ultimos)) 
    app.add_handler(CommandHandler("gasto", gasto)) 
    print("Bot iniciado...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()