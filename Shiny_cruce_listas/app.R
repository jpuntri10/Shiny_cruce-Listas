


library(shiny)
library(readxl)
library(dplyr)
library(RODBC)
library(DT)
library(writexl)

ui <- fluidPage(
  titlePanel("Cruce de Documentos con listas"),
  sidebarLayout(
    sidebarPanel(
      fileInput("archivo", "Sube tu archivo Excel con COD_DOCUM", accept = ".xlsx"),
      textInput("doc_manual", "O ingresa un documento manualmente"),
      actionButton("procesar", "Procesar")
    ),
    mainPanel(
      DTOutput("tabla_resultado"),
      br(),
      downloadButton("descargar_excel", "Descargar Resultado en Excel (.xlsx)")
    )
  )
)

server <- function(input, output, session) {
  resultado_cruce <- reactiveVal(NULL)
  
  observeEvent(input$procesar, {
    # 1) Leer códigos
    if (nzchar(input$doc_manual)) {
      codigos <- data.frame(COD_DOCUM = trimws(as.character(input$doc_manual)),
                            stringsAsFactors = FALSE)
    } else {
      req(input$archivo)
      codigos <- read_excel(input$archivo$datapath)
      if (!("COD_DOCUM" %in% names(codigos))) {
        showNotification("El Excel no tiene la columna 'COD_DOCUM'.", type = "error")
        resultado_cruce(data.frame())
        return()
      }
      codigos <- codigos[, "COD_DOCUM", drop = FALSE]
      codigos$COD_DOCUM <- trimws(as.character(codigos$COD_DOCUM))
      codigos <- codigos %>% filter(!is.na(COD_DOCUM) & COD_DOCUM != "")
      if (nrow(codigos) == 0) {
        showNotification("No hay códigos válidos para consultar.", type = "warning")
        resultado_cruce(data.frame())
        return()
      }
    }
    
    # 2) Partición para evitar IN demasiado largo
    tamaño <- 1000
    particiones <- split(codigos, ceiling(seq_len(nrow(codigos)) / tamaño))
    
    # 3) Conexión
    c <- odbcConnect("TRON_BI", uid = "jpuntri", pwd = "Peru.500600")
    on.exit(odbcClose(c), add = TRUE)
    
    cruce_listas <- data.frame()
    
    # 4) Consulta por bloques
    for (i in seq_along(particiones)) {
      cod_docum_list <- as.character(particiones[[i]]$COD_DOCUM)
      if (length(cod_docum_list) == 0) next
      cod_docum_string <- paste0("'", paste(cod_docum_list, collapse = "', '"), "'")
      
      query <- paste0("
SELECT DISTINCT TIP_DOCUM, COD_DOCUM, TARGEN, NOMBRE, MCA_INH, FEC_ACTU, DETALLE
FROM (
    -- TARGEN60
    SELECT TIP_DOCUM, COD_DOCUM, 
           CASE COD_DOCUM WHEN '' THEN 'TARGEN 60' ELSE 'TARG 60' END AS TARGEN,
           NULL AS DETALLE,
           NULL AS NOMBRE, MCA_INH, FEC_ACTU
    FROM TARGEN60
    WHERE MCA_INH <> 'S'

    UNION ALL

    -- TARGEN66
    SELECT TIP_DOCUM, COD_DOCUM, 
           CASE COD_DOCUM WHEN '' THEN 'TARGEN 66' ELSE 'TARG 66' END AS TARGEN,
           SECTOR AS DETALLE,
           NOMBRE_O_RAZON_SOCIAL AS NOMBRE, MCA_INH, FEC_ACTU
    FROM TRON2000.TARGEN66
    WHERE MCA_INH <> 'S'

    UNION ALL

    -- TARGEN73
    SELECT TIP_DOCUM, COD_DOCUM, 
           CASE COD_DOCUM WHEN '' THEN 'TARGEN 73' ELSE 'TARG 73' END AS TARGEN,
           CARGO AS DETALLE,
           NOMBRE1 || ' ' || APELLIDO_PATERNO || ' ' || APELLIDO_MATERNO AS NOMBRE, MCA_INH, FEC_ACTU
    FROM TRON2000.TARGEN73
    WHERE MCA_INH <> 'S'

    UNION ALL

    -- lcfo_lst_negra
    SELECT TIPO_ID AS TIP_DOCUM, COD_ID AS COD_DOCUM, 
           CASE COD_ID WHEN '' THEN 'LCFO' ELSE 'LCFO LST NEGRA' END AS TARGEN,
           DETA_MOT AS DETALLE,
           NOM_COMPLETO AS NOMBRE, MCA_INH, FEC_ACTU
    FROM TRON2000.lcfo_lst_negra
    WHERE MCA_INH <> 'S'
)
WHERE COD_DOCUM IN (", cod_docum_string, ")
      ")
  
  resultado <- sqlQuery(c, query, as.is = TRUE)
  if (!is.null(resultado) && nrow(resultado) > 0) {
    cruce_listas <- bind_rows(cruce_listas, resultado)
  }
    }
    
    # 5) Saneamiento de datos
    if (!is.null(cruce_listas) && nrow(cruce_listas) > 0) {
      # Factor/list -> char
      cruce_listas <- cruce_listas %>%
        mutate(across(everything(), ~ {
          if (is.list(.)) {
            sapply(., function(x) if (length(x) == 0 || is.null(x)) "" else as.character(x))
          } else if (is.factor(.)) {
            as.character(.)
          } else {
            .
          }
        }))
      
      # Texto a UTF-8, sin controles ni saltos
      cruce_listas <- cruce_listas %>%
        mutate(across(where(is.character), function(x) {
          x <- enc2utf8(x)
          x <- iconv(x, from = "latin1", to = "UTF-8//TRANSLIT", sub = "")
          x <- gsub("[\\r\\n\\t]+", " ", x, perl = TRUE)
          x <- gsub("[\\x00-\\x1F\\x7F]", "", x, perl = TRUE)
          x <- gsub(" +", " ", x, perl = TRUE)
          trimws(x)
        }))
      
      # 👉 FEC_ACTU: dejar solo FECHA (sin horas)
      if ("FEC_ACTU" %in% names(cruce_listas)) {
        if (inherits(cruce_listas$FEC_ACTU, "POSIXct")) {
          cruce_listas$FEC_ACTU <- as.Date(cruce_listas$FEC_ACTU)
        } else if (inherits(cruce_listas$FEC_ACTU, "Date")) {
          # ya es Date, lo dejamos
          cruce_listas$FEC_ACTU <- cruce_listas$FEC_ACTU
        } else {
          # Viene como texto: tomar solo la parte de fecha (antes del espacio) y convertir
          # Ejemplos: "2024-11-10 00:00:00" -> "2024-11-10"
          #           "2024/11/10 15:20"    -> "2024/11/10"
          fecha_txt <- sub(" .*", "", cruce_listas$FEC_ACTU)  # corta después del primer espacio
          # Intentar con "-" primero; si falla, intentar con "/"
          fecha_try <- suppressWarnings(as.Date(fecha_txt, format = "%Y-%m-%d"))
          idx_na <- is.na(fecha_try)
          if (any(idx_na)) {
            fecha_try[idx_na] <- suppressWarnings(as.Date(fecha_txt[idx_na], format = "%Y/%m/%d"))
          }
          # Si aún hay NA, dejar vacío para evitar problemas
          fecha_try[is.na(fecha_try)] <- as.Date(NA)
          cruce_listas$FEC_ACTU <- fecha_try
        }
      }
      
      # Reemplazar NAs seguros
      cruce_listas <- cruce_listas %>%
        mutate(across(where(function(x) !inherits(x, "Date")), ~ replace(., is.na(.), "")))
      # Nota: para Date, Excel acepta NA como celdas en blanco
    }
    
    resultado_cruce(cruce_listas)
    
    output$tabla_resultado <- renderDT({
      datatable(cruce_listas, options = list(pageLength = 10))
    })
  })

# 6) Descarga XLSX (writexl)
output$descargar_excel <- downloadHandler(
  filename = function() {
    paste0("resultado_cruce_", format(Sys.time(), "%Y%m%d_%H%M"), ".xlsx")
  },
  content = function(file) {
    df <- resultado_cruce()
    
    if (is.null(df) || nrow(df) == 0) {
      wb <- list(Resultado = data.frame(Mensaje = "Sin datos para mostrar"))
      writexl::write_xlsx(wb, path = file)
      return()
    }
    
    # Asegurar nombres de columnas
    names(df) <- make.names(names(df), unique = TRUE)
    
    # Escribir: si FEC_ACTU es Date, Excel mostrará solo fecha
    wb <- list(Resultado = df)
    writexl::write_xlsx(wb, path = file)
  }
)
}

shinyApp(ui = ui, server = server)

