# -Sistema-Automatizado-de-Gesti-n-de-Datos-en-Excel

📊 Sistema Automatizado de Gestión de Datos en Excel

🚀 Descripción del Proyecto

Este proyecto consiste en el desarrollo de un mini-sistema automatizado de gestión de datos en Microsoft Excel, diseñado para registrar información, procesarla automáticamente y realizar consultas dinámicas mediante una interfaz interactiva.

El sistema fue construido siguiendo una arquitectura de flujo de datos basada en:

Entrada de Datos → Procesamiento → Consulta y Control

La solución simula el funcionamiento de un sistema administrativo interno, integrando automatización, limpieza de datos y herramientas de consulta para optimizar el manejo de información dentro de un entorno Excel.
__________________________________________________________________________________________________________________________________________________________________________________________

🛠️ Tecnologías y Herramientas Utilizadas

-Microsoft Excel
-VBA (Visual Basic for Applications)
-Power Query
-Tablas Dinámicas
-Validación de Datos
-Formato Condicional
-Funciones avanzadas de Excel

*Funciones implementadas*
-BUSCARV
-FILTRAR
-ELEGIRCOLS / ELEGIRCOLUMNAS
-ÚNICOS
-CONTARA
-SUMA
__________________________________________________________________________________________________________________________________________________________________________________________

🏗️ Arquitectura del Mini-Sistema

El archivo fue estructurado de forma modular para separar correctamente cada proceso del sistema:

🏠 Menú Principal

Se desarrolló una hoja de inicio con navegación interactiva mediante botones y macros VBA, permitiendo acceder rápidamente a las diferentes secciones del sistema.

*Funcionalidades*
-Navegación entre hojas mediante botones
-Interfaz visual tipo sistema administrativo
-Acceso rápido a formularios y consultas
<img width="772" height="560" alt="image" src="https://github.com/user-attachments/assets/d8876b2b-f021-4ea5-8b33-98195a13a9b4" />
________________________________________________________________________________________________________________________________________________________________

📝 Formulario de Registro

Se diseñó un formulario automatizado para el ingreso de información mediante una interfaz visual estructurada.

*Características*
-Registro automatizado de datos
-Generación automática de IDs únicos
-Botones programados con VBA
-Validación de campos obligatorios
-Limpieza automática del formulario después de guardar registros
-Inserción automática de datos en la base de datos principal

<img width="645" height="526" alt="image" src="https://github.com/user-attachments/assets/defc5786-3f10-4589-81eb-f49421c715d0" />
________________________________________________________________________________________________________________________________________________________________

🗂️ Gestión de Base de Datos

El sistema almacena la información utilizando tablas estructuradas para mantener una organización adecuada de los registros.

*Implementaciones*
-Base de datos principal de registros
-Tabla independiente para observaciones
-Relación entre registros mediante ID
-Organización estructurada de información

Las observaciones de cada compra fueron almacenadas en una tabla separada que contiene:

-ID del registro
-Fecha
-Observación correspondiente

<img width="1201" height="581" alt="image" src="https://github.com/user-attachments/assets/3099923e-2769-4ea8-aaa9-09fd0655638e" />
<img width="625" height="577" alt="image" src="https://github.com/user-attachments/assets/a84ad4df-86e7-4073-910a-8b7892210a44" />
________________________________________________________________________________________________________________________________________________________________

🧹 Limpieza y Transformación de Datos

Se implementó un proceso ETL básico utilizando Power Query para automatizar la limpieza y transformación de la información registrada.

*Procesos realizados:*
-Limpieza de datos
-Transformación de formatos
-Normalización de información
-Creación de una base de datos limpia para consultas

El resultado final fue cargado en una hoja independiente denominada:

✅ BASE_LIMPIA

La cual funciona como fuente principal para las consultas dinámicas y reportes.
________________________________________________________________________________________________________________________________________________________________
🔍 Sistema de Consultas Dinámicas

Se desarrolló una interfaz de consultas utilizando funciones avanzadas de Excel para permitir búsquedas y extracción automática de información.

*Funcionalidades implementadas:*

-Consulta de registros por ID con validación de datos [lista]
-Extracción automática de información mediante BUSCARV
-Generación dinámica de listas filtradas
-Obtención automática de productos únicos
-Conteo automático de registros
-Resumen de pedidos completados
-Visualización dinámica de información

*Funciones utilizadas:*
1. BUSCARV
2. FILTRAR
3. ELEGIRCOLS / ELEGIRCOLUMNAS
4. ÚNICOS
5. CONTARA
6. SUMA

<img width="814" height="597" alt="image" src="https://github.com/user-attachments/assets/22312462-af44-4241-8a88-90a7c982b20a" />

________________________________________________________________________________________________________________________________________________________________
🎨 Diseño e Interfaz

El sistema fue diseñado con una interfaz visual enfocada en simular una aplicación administrativa dentro de Excel.

*Características visuales:*

-Navegación mediante botones VBA
-Diseño modular por secciones
-Interfaz organizada y estructurada
-Formularios personalizados
-Acceso interactivo entre hojas
***************************************************************************************************************************************************************
💡 Objetivos del Proyecto

Automatizar procesos manuales en Excel
Implementar gestión estructurada de datos
Aplicar limpieza y transformación de información
Desarrollar interfaces interactivas en Excel
Simular un sistema administrativo funcional
Aplicar automatización mediante VBA y Power Query


📌 Habilidades Aplicadas

-Automatización con VBA
-Gestión de datos
-ETL básico con Power Query
-Consultas dinámicas
-Manipulación de tablas estructuradas
-Diseño de interfaces en Excel
-Validación y limpieza de datos
****************************************************************************************************************************************************************
AUTORA : NAOMI AYOVÍ  
****************************************************************************************************************************************************************
