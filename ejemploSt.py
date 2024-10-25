import streamlit as st
import pandas as pd
from PyPDF2 import PdfMerger
import io

# Inicialización del estado de sesión
if 'vista_actual' not in st.session_state:
    st.session_state.vista_actual = 'inicio'

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
        ('Pantalla inicial', 'Consolidado PDF', 'Opción 2')
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

    elif opcion == 'Opción 2':
        st.subheader('Pagina para probar diversos widgets de Streamlit')

        # Control deslizante
        st.markdown('**Este es un control deslizante, se podria usar para seleccionar que cantidad de estudiantes se requiere filtrar**')
        valor = st.slider('Selecciona un valor:', 0, 100, 50)
        st.write('El valor seleccionado es:', valor)

        # Selector de fechas
        st.markdown('**Este es un selector de fechas, se podria usar para seleccionar una fecha de corte par alguna entrega o inicio de etapa productiva**')
        fecha = st.date_input('Elige una fecha')
        st.write('Fecha seleccionada:', fecha)

        # Crear un DataFrame
        df = pd.DataFrame({
            'Nombre': ['Juan', 'Ana', 'Luis', 'Sofía', 'Pedro'],
            'Edad': [23, 30, 21, 24, 28],
            'Ciudad': ['Bogotá', 'Medellín', 'Cali', 'Bogotá', 'Medellín']
        })

        #Añadir la opcion ver todo
        filtros_posibles = ['Ver todos'] + list(df['Ciudad'].unique())

        # Crear un filtro interactivo
        st.markdown('**Este es un filtro interactivo, se podria usar para seleccionar algun aspecto especifico que se requiera de los aprendices**')
        ciudad_filtrada = st.selectbox("Selecciona la ciudad", filtros_posibles)

        # Filtrar el DataFrame solo si no se selecciona "Ver todos"
        if ciudad_filtrada == 'Ver todos':
            df_filtrado = df
        else:
            df_filtrado = df[df['Ciudad'] == ciudad_filtrada]

        
        # Mostrar el resultado filtrado
        st.markdown('**Así se puede ver los dataframe previamente hechos directamente desde pandas**')
        st.dataframe(df_filtrado)

# Función para mostrar la vista de aprendiz
def mostrar_aprendiz():
    if st.button('⬅️ Volver atrás', key='volver_aprendiz'):
        cambiar_vista('inicio')
        st.rerun()
    
    st.title('Consolidado PDF')
    mostrar_formulario_pdf()

# Funcion para unir PDFs
def unir_pdfs(archivos_pdf):

    # Genera un objeto PdfMerger para unir los PDFs
    merger = PdfMerger()

    for archivo in archivos_pdf:
        if archivo is not None:
            pdf_bytes = io.BytesIO(archivo.getvalue())
            merger.append(pdf_bytes)
    output = io.BytesIO()
    merger.write(output)
    output.seek(0)
    return output

# Función para mostrar/crear el formulario de subida de PDFs
def mostrar_formulario_pdf():
    st.write('Suba los archivos PDF en el campo correspondiente.')
    col1, col2 = st.columns(2)

    with col1:
        fichaPdf = st.text_input("Introduzca el número de ficha del aprendiz:")
    with col2:
        nombrePdf = st.text_input("Introduzca nombre del aprendiz:")
        
    archivos = []
    documentos = [
        "F-023(final)",
        "Agencia Publica de Empleo",
        "Paz y Salvo Academico",
        "Copia del Documento de Identidad",
        "Certificación empresa",
        "Para tecnólogos: Certificación Pruebas TyT",  # Este es opcional
        "Formato de Entrega de Documentos"
    ]
    
    # Se crea un diccionario para almacenar los archivos y así tener en cuenta cuáles han sido subidos
    archivos_subidos = {}
    
    # Crear esoacio donde se mostrara el estado de los archivos, es decir cuales se han subido y si son obligatorios
    estado_archivos = st.empty()
    
    # En esta variable se almacenará el nombre del documento opcional
    documento_opcional = "Para tecnólogos: Certificación Pruebas TyT"
    
    # Crear y mostrar los campos para subir archivos
    for nombre in documentos:
        es_opcional = nombre == documento_opcional
        label = f"{nombre} {'(Opcional)' if es_opcional else '(Obligatorio)'}"
        archivo = st.file_uploader(label, type=["pdf"], key=f"upload_{nombre}")
        archivos_subidos[nombre] = archivo
        
    # Verificar documentos obligatorios
    documentos_obligatorios = [doc for doc in documentos if doc != documento_opcional]
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
                    pdf_final = unir_pdfs(archivos_para_unir)
                    # Ofrecer el archivo para descargar
                    st.download_button(
                        label="Descargar PDF Consolidado",
                        data=pdf_final,
                        file_name=f"{fichaPdf} {nombrePdf}.pdf",
                        mime="application/pdf"
                    )
                st.success('¡PDFs unidos exitosamente!')
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
