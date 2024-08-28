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

# Número de días entre bitacoras
intervalo_dias = 15

for index, row in df.iterrows():

    # Verificar si la fila está vacía
    if pd.isna(row).all():
        print("Se encontró una fila vacía, deteniendo la iteración.")
        break
    
    # Buscar las columnas que contienen los datos necesario para realizar los filtros
    # Fecha en la que se inicio la etapa productiva
    fecha_columna = [col for col in df.columns if 'Inicio_Ficha(Productiva)' in col]
    # Columna B1
    columna_b1 = [col for col in df.columns if 'B1' in col]
    # Correo del aprendiz
    correo_aprendiz = [col for col in df.columns if 'CorreoAprendiz' in col]

    # Se verifica la fecha de inicio de la etapa productiva y se convierte a tipo dateTime
    if fecha_columna:
        fecha_columna = df[fecha_columna[0]].iloc[0]
        fecha_inicio = pd.to_datetime(fecha_columna).date()
    else:
        print("No se encontró una columna de fecha adecuada")
        break

    # Iterar sobre las columnas de bitacora para verificar cual se ha enviado
    for i in range(1, 13):
        columna_bitacora = f'B{i}'

        # Calcular la fecha en la que se debe enviar la notificación
        fecha_notificacion = fecha_inicio + timedelta(days=intervalo_dias * i)

        # Verificar si subio la bitacora
        estado_bitacora = df[columna_bitacora].iloc[0].strip().lower()

        # Verificar si se debe enviar la notificación
        if estado_bitacora == 'no' and fecha_notificacion <= fecha_actual:
            # Obtener el destinatario del correo electrónico
                destinatario = df[correo_aprendiz[0]].iloc[0]

                # Obtener los detalles del correo electrónico
                if destinatario:

                    msg = MIMEMultipart()
                    msg['From'] = 'astroc2208@gmail.co'
                    msg['To'] = destinatario
                    msg['Subject'] = f'Falta entrega de bitacora {i}'
                    body = f'Por medio de este correo se le notifica que la bitacora {i} no ha sido entregada y debio haber sido subida el dia {fecha_notificacion}.'
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
                break     
        else:
            print(f'{columna_bitacora} ya fue enviada o el aprendiz tiene tiempo de enviarla.')

print("Proceso completado")