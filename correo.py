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
intervalo_dias = 14

for index, row in df.iterrows():

    # Verificar si la fila está vacía
    if pd.isna(row).all():
        print("Se encontró una fila vacía, deteniendo la iteración.")
        break
    
    # Fecha en la que se inicio la etapa productiva
    fecha_columna = [col for col in df.columns if 'Inicio_Ficha(Productiva)' in col]

    # Hallar la columna y el valor del nombre del aprendiz
    aprendiz = [col for col in df.columns if 'Aprendiz' in col]
    nombre_aprendiz = df[aprendiz[0]].iloc[index]

    # Hallar la columna y el valor del correo del aprendiz
    correo_aprendiz = [col for col in df.columns if 'CorreoAprendiz' in col]
    destinatario = df[correo_aprendiz[0]].iloc[index]

    # Hallar la columna y el valor del correo del instructor
    correo_instructor = [col for col in df.columns if 'instructor_seguimiento' in col]
    destinatario2 = df[correo_instructor[0]].iloc[index]
    print(destinatario2)

    # Hallar la columna y el valor de Bitacoras
    bitacoras = [col for col in df.columns if 'Bitacoras' in col]
    cantidad_bitacoras = df[bitacoras[0]].iloc[index]

    # Se verifica la fecha de inicio de la etapa productiva y se convierte a tipo dateTime
    if fecha_columna:
        fecha_columna = df[fecha_columna[0]].iloc[index]
        fecha_inicio = pd.to_datetime(fecha_columna).date()
    else:
        print("No se encontró una columna de fecha adecuada")
        break

    # Verificar que la fecha actual no sea mayor a la fecha de entrega de la sexta bitacora
    if fecha_actual > fecha_inicio + timedelta(days=intervalo_dias * 6) and cantidad_bitacoras < 5:

        print(nombre_aprendiz,cantidad_bitacoras)
        # Obtener los detalles del correo electrónico
        msg = MIMEMultipart()
        msg['From'] = 'astroc2208@gmail.com'
        msg['To'] = str(destinatario2)
        msg['Cc'] = str(destinatario)
        msg['Subject'] = f'Fomalización de citación a comité aprendiz {nombre_aprendiz}'
        body = f'Se le notifica que el aprendiz {nombre_aprendiz} ha subido {cantidad_bitacoras} bitacoras y ya ha pasado el tiempo de entrega de la sexta bitacora, por lo que se solicita que formalice la citación a comité'
        msg.attach(MIMEText(body, 'plain'))

        # Conexión al servidor SMTP
        smtp = smtplib.SMTP('smtp.gmail.com', 587)
        smtp.starttls()

        # Autenticación con tu correo y contraseña de aplicación
        smtp.login('astroc2208@gmail.com', 'jsgm gpyz gakh ywop')

        # Envío del correo
        smtp.send_message(msg)
        print(f"El aprendiz {nombre_aprendiz} tiene menos de 5 bitacoras subidas y ya paso el tiempo de entrega de la sexta")

        # Cierre de la conexión SMTP
        smtp.quit()
    
    elif fecha_actual > fecha_inicio + timedelta(days=intervalo_dias * 12) and cantidad_bitacoras < 12:

        print(nombre_aprendiz,cantidad_bitacoras)
        # Obtener los detalles del correo electrónico
        msg = MIMEMultipart()
        msg['From'] = 'astroc2208@gmail.com'
        msg['To'] = destinatario, destinatario2
        msg['Subject'] = 'Fomalización de citación a comité prendiz {nombre_aprendiz}'
        body = f'Se le notifica que el aprendiz {nombre_aprendiz} ha subido {cantidad_bitacoras} bitacoras y ya ha pasado el tiempo de entrega de la ultima bitacora, por lo que se solicita que formalice la citación a comité'
        msg.attach(MIMEText(body, 'plain'))

        # Conexión al servidor SMTP
        smtp = smtplib.SMTP('smtp.gmail.com', 587)
        smtp.starttls()

        # Autenticación con tu correo y contraseña de aplicación
        smtp.login('astroc2208@gmail.com', 'jsgm gpyz gakh ywop')

        # Envío del correo
        smtp.send_message(msg)
        print(f"El aprendiz {nombre_aprendiz} tiene menos de 5 bitacoras subidas y ya paso el tiempo de entrega de la sexta")

        # Cierre de la conexión SMTP
        smtp.quit()

    else:
        # Se verifica el nombre del aprendiz sobre el que se va a realizar la revisión
        print(nombre_aprendiz)
        # Iteración sobre las columnas de bitacora para verificar cual se ha enviado
        for i in range(1, 13):

            # Nombre de la columna de la bitacora que se esta revisando
            columna_bitacora = f'B{i}'

            # Calcular la fecha en la que se debe enviar la notificación
            fecha_notificacion = fecha_inicio + timedelta(days=intervalo_dias * i)

            # Verificar si subio la bitacora o no
            estado_bitacora = df[columna_bitacora].iloc[index].strip().lower()

            # Verificar si se debe enviar la notificación, comprando la fecha en la que debio subir la bitacora con la fecha actual y si la bitacora no ha sido enviada
            if estado_bitacora == 'no' and fecha_notificacion <= fecha_actual:

                # Verificar si se tiene un correo electrónico registrado
                if destinatario:

                    # Obtener los detalles del correo electrónico
                    msg = MIMEMultipart()
                    msg['From'] = 'astroc2208@gmail.com'
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
                    print(f"Falta bitacora {i} que debia entregarce el {fecha_notificacion}, se ha enviado notificación a {destinatario}")

                    # Cierre de la conexión SMTP
                    smtp.quit()

                else:
                    print("No se encontró un correo electronico registrado")  
                break
            else:
                print(f'{columna_bitacora} ya fue enviada o el aprendiz tiene tiempo de enviarla hasta {fecha_notificacion}.')

print("Proceso completado")