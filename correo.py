import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dateutil.relativedelta import relativedelta

# Leer la base de datos de Excel
df = pd.read_excel(r'C:/Sena/Programacion de Software/Etapa productiva/requerimientosSenaBeta/requerimientosSenaBeta/Sources/prototiposSeguimiento.xlsx')

#limpiar el nombre de las columnas
df.columns = df.columns.str.strip()

# Obtener la fecha actual
fecha_actual = datetime.now().date()

# Número de días entre bitacoras
intervalo_dias = 14


def requerimiento_1():

    # Convertir la columna de fecha en un formato datetime
    df['Inicio_Ficha(Productiva)'] = pd.to_datetime(df['Inicio_Ficha(Productiva)'])

    # Obtener la fecha actual
    fecha_actual = pd.to_datetime('today')

    # Filtrar fichas que han pasado 12 y 18 meses
    fichas_12_meses = df[df['Inicio_Ficha(Productiva)'] + relativedelta(months=12) <= fecha_actual]
    fichas_18_meses = df[df['Inicio_Ficha(Productiva)'] + relativedelta(months=18) <= fecha_actual]

    # Eliminar duplicados basados en la columna 'Ficha'
    fichas_12_meses = fichas_12_meses.drop_duplicates(subset=['Ficha'])
    fichas_18_meses = fichas_18_meses.drop_duplicates(subset=['Ficha'])

    # Solo guardar la columna 'Ficha'
    fichas_12_meses = fichas_12_meses[['Ficha']]
    fichas_18_meses = fichas_18_meses[['Ficha']]

    # Guardar en archivos Excel separados
    fichas_12_meses.to_excel('fichas_12_meses.xlsx', index=False)
    fichas_18_meses.to_excel('fichas_18_meses.xlsx', index=False)

def requerimientos_2_3_4():
                
        if alternativa != 'NA':

            #  Se envia notificacion si no se ha entregado acta de inicio antes de una semana
            if fecha_actual < fecha_inicio + timedelta(days=7) and acta_inicio_valor == 'NO':

                #se envia un correo al aprendiz
                enviar_correo_aprendiz(f'Flata acta de inicio por entregar', f'Por medio de este correo se le recuerda que no ha entregado acta de inicio, se le recuerda que si no la entrega antres de {fecha_inicio + timedelta(days=7)} se formalizara la citación a comite.')
                print(f'Falta acta de inicio por entregar, se ha enviado notificación a {destinatario}')

            elif fecha_actual > fecha_inicio + timedelta(days=7) and acta_inicio_valor == 'NO':

                # Se modifica la columna COMITÉ por pendiente para que el instructor pueda ver que se debe formalizar la citación
                df.at[index, 'COMITÉ'] = 'pendiente'

                #se envia un correo al instructor
                enviar_correo_instructor(f'Fomalización de citación a comité aprendiz {nombre_aprendiz}', f'Se le notifica que el aprendiz {nombre_aprendiz} no ha entregado el acta de inicio y ya ha pasado el tiempo de entrega, por lo que se solicita que formalice la citación a comité.')
                print(f'Falta acta de inicio por entregar y ya ha pasado una semana, se ha enviado notificación de formalizacion a comite a {destinatario2}')

            else:
                # Verificar que la fecha actual no sea mayor a la fecha de entrega de la sexta bitacora
                if fecha_actual > fecha_inicio + timedelta(days=intervalo_dias * 6) and cantidad_bitacoras < 5:

                    # Se modifica la columna COMITÉ por pendiente para que el instructor pueda ver que se debe formalizar la citación
                    df.at[index, 'COMITÉ'] = 'pendiente'

                    # Enviar correo al instructor con copia al aprendiz
                    enviar_correo_instructor(f'Fomalización de citación a comité aprendiz {nombre_aprendiz}', f'Se le notifica que el aprendiz {nombre_aprendiz} ha subido {cantidad_bitacoras} bitacoras y ya ha pasado el tiempo de entrega de la sexta bitacora, por lo que se solicita que formalice la citación a comité.')
                
                elif fecha_actual > fecha_inicio + timedelta(days=intervalo_dias * 12) and cantidad_bitacoras < 12:

                    # Se modifica la columna COMITÉ por pendiente para que el instructor pueda ver que se debe formalizar la citación
                    df.at[index, 'COMITÉ'] = 'pendiente'

                    # Enviar correo al instructor con copia al aprendiz
                    enviar_correo_instructor(f'Fomalización de citación a comité aprendiz {nombre_aprendiz}', f'Se le notifica que el aprendiz {nombre_aprendiz} ha subido {cantidad_bitacoras} bitacoras y ya ha pasado el tiempo de entrega de la bitacora 12, por lo que se solicita que formalice la citación a comité.')

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

                                # Se envia correo al aprendiz
                                enviar_correo_aprendiz(f'Falta entrega de bitacora {i}', f'Por medio de este correo se le notifica que la bitacora {i} no ha sido entregada y debio haber sido subida el dia {fecha_notificacion}.')
                                print(f"Falta bitacora {i} que debia entregarce el {fecha_notificacion}, se ha enviado notificación a {destinatario}")

                            else:
                                print("No se encontró un correo electronico registrado")  
                            break
                        else:
                            print(f'{columna_bitacora} ya fue enviada o el aprendiz tiene tiempo de enviarla hasta {fecha_notificacion}.')
        else:
            # En esta parte se verificaria las fechas 2 y 3, es decir a los 12 y 18 meses de le fecha de inicio etapa productiva.
            fecha_12_meses = fecha_inicio + relativedelta(months=12)
            print(fecha_12_meses)

            if fecha_actual > fecha_12_meses:
                enviar_correo_aprendiz(f'No se ha elegido etapa productiva en el tiempo establecido', f'Por medio de este correo se le notifica que el dia {fecha_12_meses} se acabo el plazo para elegir una alternativa de etapa productiva, pues ya han pasado 12 meses desde la finalización de etapa lectiva y ya no le queda plazo de completarla, por lo que se iniciara proceso de deserción, si desea revertir esto haga envio de evidencias a su instructor de seguimiento al correo: {destinatario2}.')
                print(f'No se ha elegido etapa productiva en el tiempo establecido, se ha enviado notificación a {destinatario}')

            else:
                # Se verifica cuando fue la ultima vez que se envio correo, si ya cumple los 15 dias sin elegir etapa productiva se envia correo
                if pd.isna(fecha_envio) and fecha_actual > fecha_inicio + timedelta(days=14):
                    enviar_correo_aprendiz(f'Recordatorio alternativa etapa productiva', f'Por medio de este correo se le recuerda que debe elegir una alternativa de etapa productiva antes de {fecha_12_meses}, de otro modo no alcanzara a completar la misma y se le iniciara proceso de deserción.')
                    print(f'Recordatorio alternativa etapa productiva, se ha enviado notificación a {destinatario}')
                    df.at[index, 'Fecha_Envio'] = fecha_actual

                elif fecha_actual > fecha_envio + timedelta(days=14):
                    fecha_envio = fecha_actual
                    enviar_correo_aprendiz(f'Recordatorio alternativa etapa productiva', f'Por medio de este correo se le recuerda que debe elegir una alternativa de etapa productiva antes de {fecha_12_meses}, de otro modo no alcanzara a completar la misma y se le iniciara proceso de deserción.')
                    print(f'Recordatorio alternativa etapa productiva, se ha enviado notificación a {destinatario}')
                    df.at[index, 'Fecha_Envio'] = fecha_actual

        # Guardar los cambios en el archivo de Excel
        df.to_excel('C:/Sena/Programacion de Software/Etapa productiva/requerimientosSenaBeta/requerimientosSenaBeta/Sources/prototiposSeguimiento.xlsx', index=False)

# Funcion para enviar correo al instructor con copia al aprendiz
def enviar_correo_instructor(asunto, cuerpo):

    # Obtener los detalles del correo electrónico
        msg = MIMEMultipart()
        msg['From'] = 'astroc2208@gmail.com'
        msg['To'] = str(destinatario2)
        msg['Cc'] = str(destinatario)
        msg['Subject'] = asunto
        body = cuerpo
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

# funcion para enviar correo al aprendiz
def enviar_correo_aprendiz(asunto, cuerpo):

    # Obtener los detalles del correo electrónico
    msg = MIMEMultipart()
    msg['From'] = 'astroc2208@gmail.com'
    msg['To'] = destinatario
    msg['Subject'] = asunto
    body = cuerpo
    msg.attach(MIMEText(body, 'plain'))

    # Conexión al servidor SMTP
    smtp = smtplib.SMTP('smtp.gmail.com', 587)
    smtp.starttls()

    # Autenticación con tu correo y contraseña de aplicación
    smtp.login('astroc2208@gmail.com', 'jsgm gpyz gakh ywop')

    # Envío del correo
    smtp.send_message(msg)

    # Cierre de la conexión SMTP
    smtp.quit()

# llamar a la función de requerimiento 1, solo se necesita una vez por eso esta fuera del ciclo
# requerimiento_1()

# Iterar sobre las filas del DataFrame, esto se realiza ya que necesitamos recorrer cada fila para verificar si se cumple con los requerimientos
for index, row in df.iterrows():

    # Verificar si la fila está vacía
    if pd.isna(row).all():
        print("Se encontró una fila vacía, deteniendo la iteración.")
        break
    

    # Se verifica si aprendiz eligio alternativa etapa productiva
    etapa_productiva = [col for col in df.columns if 'Alternativa(Equipo Etapa Productiva)' in col]
    alternativa = df[etapa_productiva[0]].iloc[index]    

    # Hallar la columna y el valor del nombre del aprendiz
    aprendiz = [col for col in df.columns if 'Aprendiz' in col]
    nombre_aprendiz = df[aprendiz[0]].iloc[index]

    # Hallar la columna y el valor del correo del aprendiz
    correo_aprendiz = [col for col in df.columns if 'CorreoAprendiz' in col]
    destinatario = df[correo_aprendiz[0]].iloc[index]

    # Hallar la columna y el valor del correo del instructor
    correo_instructor = [col for col in df.columns if 'instructor_seguimiento' in col]
    destinatario2 = df[correo_instructor[0]].iloc[index]

    # Hallar columna y valor de formato acta de inicio
    acta_inicio = [col for col in df.columns if 'ActaInicio' in col]
    acta_inicio_valor = df[acta_inicio[0]].iloc[index]

    # Hallar la columna y el valor de Bitacoras
    bitacoras = [col for col in df.columns if 'Bitacoras' in col]
    cantidad_bitacoras = df[bitacoras[0]].iloc[index]

    # Se saca la ultima fecha en que se envio correo
    fecha_envio = row['correo_verificacion']

    # Fecha en la que se inicio la etapa productiva
    fecha_columna = [col for col in df.columns if 'Inicio_Ficha(Productiva)' in col]

    # Se verifica la fecha de inicio de la etapa productiva y se convierte a tipo dateTime
    if fecha_columna:
        fecha_columna = df[fecha_columna[0]].iloc[index]
        fecha_inicio = pd.to_datetime(fecha_columna).date()
    else:
        print("No se encontró una columna de fecha adecuada")
        break

    # Llamar a la función de requerimientos 2, 3 y 4, se necesita para cada aprendiz por esto esta dentro del ciclo
    requerimientos_2_3_4()
