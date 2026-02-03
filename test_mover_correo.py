import imaplib
import logging

# --- CONFIGURACIÓN (Ajusta estos datos) ---
EMAIL = "alertasflm@indra.es"
PASSWORD = "es8EaB63"
SERVER = "imap.indra.es"
PORT_EMAIL = 993
CARPETA_ORIGEN = "BOT2-PENDIENTES"
CARPETA_DESTINO = "BOT2-PROCESADOS"

def test_mover_correo():
    mail = None
    try:
        # 1. Conexión y Login
        print(f"Intentando conectar a {SERVER}...")
        mail = imaplib.IMAP4_SSL(SERVER, PORT_EMAIL)
        mail.login(EMAIL, PASSWORD)
        print("✅ Login exitoso.")

        # 2. Seleccionar carpeta de origen
        status, _ = mail.select(CARPETA_ORIGEN)
        if status != 'OK':
            print(f"❌ Error: No se pudo encontrar la carpeta '{CARPETA_ORIGEN}'")
            return

        # 3. Buscar el correo más reciente
        status, data = mail.search(None, "ALL")
        mail_ids = data[0].split()

        if not mail_ids:
            print(f"⚠️ No hay correos para mover en '{CARPETA_ORIGEN}'.")
            return

        # Tomamos el último ID (el más reciente)
        msg_id = mail_ids[-1].decode()
        print(f"🔎 Correo encontrado. ID: {msg_id}. Intentando mover a '{CARPETA_DESTINO}'...")

        # 4. Intentar mover (Método COPY + DELETE por ser el más compatible)
        # Paso A: Copiar
        print(f"Paso A: Copiando mensaje {msg_id}...")
        result, data = mail.copy(msg_id, CARPETA_DESTINO)
        
        if result == 'OK':
            print(f"✅ Copia exitosa en '{CARPETA_DESTINO}'.")
            
            # Paso B: Marcar para borrar en origen
            print("Paso B: Marcando original como eliminado...")
            mail.store(msg_id, '+FLAGS', '\\Deleted')
            
            # Paso C: Borrar físicamente
            print("Paso C: Ejecutando EXPUNGE...")
            mail.expunge()
            print("✨ ¡Proceso completado! El correo debería haberse movido.")
        else:
            print(f"❌ Error al copiar: {result} {data}")

    except Exception as e:
        print(f"💥 Error crítico durante el test: {str(e)}")
    
    finally:
        if mail:
            try:
                mail.logout()
                print("🔒 Conexión cerrada.")
            except:
                pass

if __name__ == "__main__":
    test_mover_correo()