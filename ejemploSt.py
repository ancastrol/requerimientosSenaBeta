import streamlit as st
import pandas as pd
from PyPDF2 import PdfMerger
import io
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from dateutil.relativedelta import relativedelta
from docx import Document
import os
from zipfile import ZipFile
import tempfile
from tqdm import tqdm

# Leer excel prueba 
df = pd.read_excel(r'Sources/prototiposSeguimiento.xlsx')

# Inicialización del estado de sesión
if 'vista_actual' not in st.session_state:
    st.session_state.vista_actual = 'inicio'

# Recibe los documentos excel y word y regresa un zip con los documentos word llenos
def procesar_documentos(df, plantilla_word):

    # Crear un archivo ZIP en memoria
    zip_archivo = io.BytesIO()
    
    with ZipFile(zip_archivo, 'w') as zip_file:

        # Contador de documentos procesados
        documentos_exitosos = 0
        errores = []
        
        # Usar plantilla desde bytes
        plantilla_temp = io.BytesIO(plantilla_word)
        
        # Procesar cada aprendiz
        for index, aprendiz in df.iterrows():
            try:
                # Cargar la plantilla para cada aprendiz
                doc = Document(plantilla_temp)
                plantilla_temp.seek(0)  # Pone el "puntero" al inicio del archivo
                
                # Reemplazar los marcadores en el documento
                for paragraph in doc.paragraphs:
                    for run in paragraph.runs:
                        texto = run.text
                        for columna in df.columns:
                            marcador = '{' + columna.upper() + '}'
                            if marcador in texto:
                                texto = texto.replace(marcador, str(aprendiz[columna]))
                        run.text = texto
                
                # También buscar en las tablas
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for paragraph in cell.paragraphs:
                                for run in paragraph.runs:
                                    texto = run.text
                                    for columna in df.columns:
                                        marcador = '{' + columna.upper() + '}'
                                        if marcador in texto:
                                            texto = texto.replace(marcador, str(aprendiz[columna]))
                                    run.text = texto
                
                # Guardar documento en memoria
                doc_temp = io.BytesIO()
                doc.save(doc_temp)
                doc_temp.seek(0)
                
                # Generar nombre único para el archivo
                nombre_archivo = f"{aprendiz['Nombre']}_{aprendiz['Apellidos']}_{'acta de inicio'}.docx"
                
                # Añadir al ZIP
                zip_file.writestr(nombre_archivo, doc_temp.getvalue())
                documentos_exitosos += 1
                
            except Exception as e:
                errores.append(f"Error procesando aprendiz {aprendiz['Nombre']} {aprendiz['Apellidos']}: {str(e)}")
    
    return zip_archivo, documentos_exitosos, errores

# Función para cambiar la vista actual de la aplicación
def cambiar_vista(nueva_vista):
    st.session_state.vista_actual = nueva_vista

# Función para mostrar la vista de inicio  
def mostrar_inicio():
    st.title('Elija su rol para continuar')
    st.subheader('Seleccione una de las siguientes opciones para continuar con las herramientas indicadas para su rol.')
    st.write('')
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button('Instructor', key='btn_instructor'):
            cambiar_vista('instructor')
            st.rerun()
            
    with col2:
        if st.button('Aprendíz', key='btn_aprendiz'):
            cambiar_vista('aprendiz')
            st.rerun()

# Función para mostrar la vista de instructor
def mostrar_instructor():
    if st.button('⬅️ Volver atrás', key='volver_instructor'):
        cambiar_vista('inicio')
        st.rerun()
        
    # Crear la barra lateral
    st.sidebar.title("Menú desplegable")
    opcion = st.sidebar.selectbox(
        'Elige una opción:',
        ('Pantalla inicial', 'Consolidado PDF', 'Cruce de correspondencia')
    )
    
    if opcion == 'Pantalla inicial':

        col1, col2 = st.columns(2)

        with col1:
            st.title('Herramientas de Desarrollo Etapa Productiva')

        with col2:
            # Mostrar la imagen en la ventana
            st.image("picture103.jpg", width=200)

        # Contenido de la pagina
        st.write('Este aplicativo busca facilitar multiples tareas de los instructores con respecto al manejo de sus aprendices que estan terminando la etapa productiva y estan en curso de certificarse. Si desea ver las funcionalidades disponibles se encuentra en la barra lateral a la izquierda de la pantalla.')

    elif opcion == 'Consolidado PDF':
        mostrar_formulario_pdf()

    elif opcion == 'Cruce de correspondencia':

        st.title("🎓 Generador de Documentos para Aprendices")
        
        # Agregar información de uso
        with st.expander("ℹ️ Instrucciones de uso"):
            st.markdown("""
            1. **Preparación del Excel:**
            * Asegúrate de que tu Excel tenga encabezados claros
            * Cada columna debe corresponder a un marcador en la plantilla Word
            
            2. **Preparación de la Plantilla Word:**
            * Usa marcadores entre llaves, ejemplo: {NOMBRE}, {APELLIDO}
            * Los marcadores deben coincidir con los nombres de las columnas del Excel
            
            3. **Proceso:**
            * Sube tu archivo Excel con los datos
            * Sube tu plantilla Word
            * Haz clic en 'Generar Documentos'
            * Descarga el archivo ZIP con todos los documentos generados
            """)
        
        # Columnas para la carga de archivos
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📊 Archivo Excel")
            excel_file = st.file_uploader("Sube tu archivo Excel", type=['xlsx', 'xls'])
            
            if excel_file is not None:
                try:
                    df = pd.read_excel(excel_file)
                    st.success(f"✅ Excel cargado exitosamente - {len(df)} registros encontrados")
                    
                    # Mostrar vista previa de los datos
                    with st.expander("👀 Vista previa de datos"):
                        st.dataframe(df.head())
                    
                    # Mostrar los marcadores disponibles
                    st.info("🔍 Marcadores disponibles:")
                    marcadores = [f"{{{col.upper()}}}" for col in df.columns]
                    st.code(", ".join(marcadores))
                except Exception as e:
                    st.error(f"❌ Error al leer el archivo Excel: {str(e)}")
                    df = None
        
        with col2:
            st.subheader("📄 Plantilla Word")
            word_file = st.file_uploader("Sube tu plantilla Word", type=['docx'])
            
            if word_file is not None:
                st.success("✅ Plantilla Word cargada exitosamente")
        
        # Botón para generar documentos
        if st.button("🚀 Generar Documentos", disabled=(excel_file is None or word_file is None)):
            if excel_file is not None and word_file is not None:
                with st.spinner("Generando documentos..."):
                    try:
                        # Procesar documentos
                        zip_archivo, docs_exitosos, errores = procesar_documentos(df, word_file.getvalue())
                        
                        # Mostrar resultados
                        st.success(f"✅ Proceso completado - {docs_exitosos} documentos generados")
                        
                        # Si hay errores, mostrarlos
                        if errores:
                            with st.expander("⚠️ Errores encontrados"):
                                for error in errores:
                                    st.error(error)
                        
                        # Botón de descarga
                        st.download_button(
                            label="📥 Descargar Documentos (ZIP)",
                            data=zip_archivo.getvalue(),
                            file_name="documentos_generados.zip",
                            mime="application/zip"
                        )
                        
                    except Exception as e:
                        st.error(f"❌ Error durante la generación: {str(e)}")

# Función para mostrar la vista de aprendiz
def mostrar_aprendiz():
    if st.button('⬅️ Volver atrás', key='volver_aprendiz'):
        cambiar_vista('inicio')
        st.rerun()
    
    st.title('Consolidado PDF')
    mostrar_formulario_pdf()


# Función para enviar correo al instructor con copia al aprendiz y adjunto
def enviar_correo_instructor(asunto, cuerpo, archivo_pdf, nombre_archivo):

    # Obtener el destinatarios del correo electrónico
    destinatario = df['instructor_seguimiento'].iloc[0]
    destinatario2 = df['CorreoAprendiz'].iloc[0]

    # Obtener los detalles del correo electrónico
    msg = MIMEMultipart()
    msg['From'] = 'astroc2208@gmail.com'
    msg['To'] = str(destinatario)
    msg['Cc'] = str(destinatario2)
    msg['Subject'] = asunto
    msg.attach(MIMEText(cuerpo, 'plain'))

    # Adjuntar el archivo PDF
    adjunto = MIMEApplication(archivo_pdf.getvalue(), _subtype="pdf")
    adjunto.add_header('Content-Disposition', 'attachment', filename = nombre_archivo)
    msg.attach(adjunto)

    # Conexión al servidor SMTP
    smtp = smtplib.SMTP('smtp.gmail.com', 587)
    smtp.starttls()

    # Autenticación con tu correo y contraseña de aplicación
    smtp.login('astroc2208@gmail.com', 'jsgm gpyz gakh ywop')

    # Envío del correo
    smtp.send_message(msg)

    # Cierre de la conexión SMTP
    smtp.quit()

# Función para verificar si la ficha se encuentra y si corresponde a un tecnólogo
def verificar_ficha_tecnologo(df, numero_ficha):
    # Verificar si el DataFrame no esta vacio
    if df is None:
        return False
    
    # Convertir la columna de ficha a string para comparación
    df['Ficha'] = df['Ficha'].astype(str)
    
    # Buscar la ficha en el DataFrame
    ficha_encontrada = df[df['Ficha'] == str(numero_ficha)]
    
    if len(ficha_encontrada) > 0:
        # Devolver el valor de Es_Tecnologo
        return ficha_encontrada['Tipo'].iloc[0]
    else:
        st.warning(f"Ficha {numero_ficha} no encontrada en el registro")
        return False
    
# Funcion para unir PDFs, tiene en cuenta que el TyT de haberlo va antes del ultimo documento
def unir_pdfs_con_orden(archivos_pdf, es_tecnologo):
    merger = PdfMerger()
    
    # Si es tecnólogo, separar el documento de tecnólogo
    doc_tecnologo = None
    otros_docs = []
    
    for archivo in archivos_pdf:
        if archivo is not None:
            # Si es el documento de tecnólogo, se guarda aparte
            if "Certificación Pruebas TyT" in str(archivo.name):
                doc_tecnologo = archivo
            else:
                otros_docs.append(archivo)
    
    # Si es tecnólogo y hay más de un documento
    if es_tecnologo and len(otros_docs) > 1:
        # Unir todos los documentos excepto el último
        for archivo in otros_docs[:-1]:
            pdf_bytes = io.BytesIO(archivo.getvalue())
            merger.append(pdf_bytes)
        
        # Insertar documento de tecnólogo antes del último
        if doc_tecnologo:
            pdf_bytes_tecnologo = io.BytesIO(doc_tecnologo.getvalue())
            merger.append(pdf_bytes_tecnologo)
        
        # Agregar el último documento
        pdf_bytes_ultimo = io.BytesIO(otros_docs[-1].getvalue())
        merger.append(pdf_bytes_ultimo)
    else:
        # Si no es tecnólogo o hay pocos documentos, unir normalmente
        for archivo in otros_docs:
            pdf_bytes = io.BytesIO(archivo.getvalue())
            merger.append(pdf_bytes)
        
        # Agregar documento de tecnólogo al final si existe
        if es_tecnologo and doc_tecnologo:
            pdf_bytes_tecnologo = io.BytesIO(doc_tecnologo.getvalue())
            merger.append(pdf_bytes_tecnologo)
    
    # Se crea el PDF final
    output = io.BytesIO()
    merger.write(output)
    output.seek(0)
    return output

# Función para mostrar/crear el formulario de subida de PDFs
def mostrar_formulario_pdf():
    st.write('Suba los archivos PDF en el campo correspondiente, al oprimir el boton "Consolidar PDFs" se uniran los archivos en un solo PDF y se enviará al instructor.')
    col1, col2 = st.columns(2)

    with col1:
        fichaPdf = st.text_input("Introduzca el número de ficha del aprendiz:")
    with col2:
        nombrePdf = st.text_input("Introduzca nombre del aprendiz:")
    
    # Verificar si es tecnólogo
    es_tecnologo = verificar_ficha_tecnologo(df, fichaPdf) == 'Tecnólogo'
        
    archivos = []
        # Documentos base
    documentos_base = [
        "F-023(final)",
        "Agencia Publica de Empleo",
        "Paz y Salvo Academico",
        "Copia del Documento de Identidad",
        "Certificación empresa",
        "Formato de Entrega de Documentos"
    ]
    
    # Definir documentos a mostrar dinámicamente
    documentos = documentos_base.copy()
    
    # Agregar o quitar el documento opcional según la selección
    if es_tecnologo:
        documentos.append("Certificación Pruebas TyT")
    
    # Se crea un diccionario para almacenar los archivos y así tener en cuenta cuáles han sido subidos
    archivos_subidos = {}
    
    # Crear esoacio donde se mostrara el estado de los archivos, es decir cuales se han subido y si son obligatorios
    estado_archivos = st.empty()
    
    # Crear y mostrar los campos para subir archivos
    for nombre in documentos:
        archivo = st.file_uploader(nombre, type=["pdf"], key=f"upload_{nombre}")
        archivos_subidos[nombre] = archivo
        
    # Verificar documentos obligatorios
    documentos_obligatorios = [doc for doc in documentos]
    documentos_faltantes = [doc for doc in documentos_obligatorios if archivos_subidos[doc] is None]
    
    # Actualizar el estado de los archivos
    if documentos_faltantes:
        estado_archivos.warning(f"Documentos obligatorios faltantes: {', '.join(documentos_faltantes)}")
    else:
        estado_archivos.success("Todos los documentos obligatorios han sido subidos")
    
    # Preparar lista de archivos para unir
    archivos_para_unir = [archivo for archivo in archivos_subidos.values() if archivo is not None]
    
    if st.button('Consolidar PDFs', key='btn_consolidar'):
        if len(documentos_faltantes) > 0:
            st.error("Debe subir todos los documentos obligatorios antes de consolidar")
        elif len(archivos_para_unir) > 0:
            try:
                # Mostrar barra de progreso
                with st.spinner('Uniendo PDFs...'):
                    # Unir los PDFs
                    pdf_final = unir_pdfs_con_orden(archivos_para_unir, es_tecnologo)
                    nombre_archivo = f"{fichaPdf} {nombrePdf}.pdf"
                    enviar_correo_instructor(f'Consolidado Aprendiz {nombrePdf}', f'El aprendiz {nombrePdf} ha subido los documentos requeridos para la finalización de la etapa productiva.', pdf_final, nombre_archivo)
                    # Ofrecer el archivo para descargar
                    st.download_button(
                        label="Descargar PDF Consolidado",
                        data=pdf_final,
                        file_name = nombre_archivo,
                        mime="application/pdf"
                    )
                st.success('El archivo consolidado se ha enviado al instructor exitosamente')
            except Exception as e:
                st.error(f'Error al unir los PDFs: {str(e)}')
        else:
            st.warning('Por favor, suba al menos un archivo PDF.')

def validar_pdf(archivo):
    try:
        if archivo is not None:
            # Verificar el tipo de archivo
            if not archivo.type == "application/pdf":
                return False
            # Intentar leer el PDF para verificar que no está corrupto
            PdfMerger().append(io.BytesIO(archivo.getvalue()))
            return True
    except:
        return False
    return False

# Lógica principal de la aplicación
def main():
    if st.session_state.vista_actual == 'inicio':
        mostrar_inicio()
    elif st.session_state.vista_actual == 'instructor':
        mostrar_instructor()
    elif st.session_state.vista_actual == 'aprendiz':
        mostrar_aprendiz()

# Iniciar la aplicación, de esta forma se evita que si se importa en otro programa no se inicie directamente
if __name__ == "__main__":
    main()
