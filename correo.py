import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Leer la base de datos de Excel
df = pd.read_excel(r'C:/Sena/Programacion de Software/Etapa productiva/requerimientosSenaBeta/requerimientosSenaBeta/Sources/prototiposSeguimiento.xlsx')

#limpiar el nombre de las columnas
df.columns = df.columns.str.strip()

# Obtener la fecha actual
fecha_actual = datetime.now().date()

# Número de días después de los cuales se debe enviar el correo electrónico
dias_umbral = 15

for index, row in df.iterrows():

    # Verificar si la fila está vacía
    if pd.isna(row).all():
        print("Se encontró una fila vacía, deteniendo la iteración.")
        break

    # Imprimir los valores de la fila actual
    print(f"Procesando fila {index}:")
    
    # Buscar la columna que contiene la fecha de inicio
    fecha_columna = [col for col in df.columns if 'Inicio_Ficha(Productiva)' in col]
    columna_b1 = [col for col in df.columns if 'B1' in col]
    correo_aprendiz = [col for col in df.columns if 'CorreoAprendiz' in col]
    print(f"Columnas encontradas: {correo_aprendiz}")

    if fecha_columna:
        fecha_columna = df[fecha_columna[0]].iloc[0]
        print(f"Columna de fecha encontrada: {fecha_columna}")
        fecha_inicio = pd.to_datetime(fecha_columna).date()

    else:
        print("No se encontró una columna de fecha adecuada")
        break
    
    # Calcular la diferencia en días entre la fecha actual y la fecha en la base de datos
    diferencia_dias = (fecha_actual - fecha_inicio).days
    print(f"Diferencia en días: {diferencia_dias}")
    
    # Verificar si han pasado al menos los días de umbral
    if diferencia_dias >= dias_umbral:
        
        # Sacar el valor de la columna B1
        valor_b1 = df[columna_b1[0]].iloc[0]

        # Verificar si la casilla específica tiene el valor "si"
        if valor_b1 != 'si':

            # Obtener el destinatario del correo electrónico
            destinatario = df[correo_aprendiz[0]].iloc[0]

            # Obtener los detalles del correo electrónico
            if destinatario:

                msg = MIMEMultipart()
                msg['From'] = 'astroc2208@gmail.co'
                msg['To'] = destinatario
                msg['Subject'] = 'Falta entrega de bitacora 1'
                # subject = 'Notificación automática'
                body = 'Este es un correo electrónico automático enviado para notificarle que ya paso la fecha y no ha hecho entrega de la bitacora #1.'
                msg.attach(MIMEText(body, 'plain'))

                # Conexión al servidor SMTP
                smtp = smtplib.SMTP('smtp.gmail.com', 587)
                smtp.starttls()

                # Autenticación con tu correo y contraseña de aplicación
                smtp.login('astroc2208@gmail.com', 'jsgm gpyz gakh ywop')

                # Envío del correo
                smtp.send_message(msg)
                print(f"Correo enviado a {destinatario}")

                # Cierre de la conexión SMTP
                smtp.quit()
                
            else:
                print("No se encontró un correo electronico registrado")
        else:
            print(f"Bitacora 1 esta subida, no es necesario el envió de la notificación")
    else:
        print(f"Han pasado {diferencia_dias} días, no se ha superado el umbral de {dias_umbral} días")

print("Proceso completado")