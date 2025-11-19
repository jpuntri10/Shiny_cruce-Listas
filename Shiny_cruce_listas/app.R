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
      downloadButton("descargar_excel", "Descargar Resultado en Excel")
    )
  )
)

server <- function(input, output) {
  resultado_cruce <- reactiveVal(data.frame())
  
  observeEvent(input$procesar, {
    codigos <- NULL
    
    # Si se ingresó un documento manual
    if (nzchar(input$doc_manual)) {
      codigos <- data.frame(COD_DOCUM = input$doc_manual)
    } else {
      req(input$archivo)
      codigos <- read_excel(input$archivo$datapath)
      codigos <- codigos[, , drop = FALSE]
    }
    
    tamaño <- 1000
    particiones <- split(codigos, ceiling(seq_len(nrow(codigos)) / tamaño))
    
    c <- odbcConnect("TRON_BI", uid = "jpuntri", pwd = "Peru.3040")
    cruce_listas <- data.frame()
    
    for (i in seq_along(particiones)) {
      cod_docum_list <- as.character(particiones[[i]]$COD_DOCUM)
      cod_docum_string <- paste0("'", paste(cod_docum_list, collapse = "', '"), "'")
      
      query <- paste0("
SELECT DISTINCT TIP_DOCUM, COD_DOCUM, TARGEN,NOMBRE,MCA_INH,FEC_ACTU, DETALLE
FROM (
    -- TARGEN60
    SELECT TIP_DOCUM, COD_DOCUM, 
           CASE COD_DOCUM WHEN '' THEN 'TARGEN 60' ELSE 'TARG 60' END AS TARGEN,
           NULL AS DETALLE,
           NULL AS NOMBRE,MCA_INH,FEC_ACTU
    FROM TARGEN60
    WHERE MCA_INH <> 'S'

    UNION ALL

    -- TARGEN66
    SELECT TIP_DOCUM, COD_DOCUM, 
           CASE COD_DOCUM WHEN '' THEN 'TARGEN 66' ELSE 'TARG 66' END AS TARGEN,
           OBSERVACIONES AS DETALLE,
           NOMBRE_O_RAZON_SOCIAL AS NOMBRE,MCA_INH,FEC_ACTU
    FROM TRON2000.TARGEN66
    WHERE MCA_INH <> 'S'

    UNION ALL

    -- TARGEN73
    SELECT TIP_DOCUM, COD_DOCUM, 
           CASE COD_DOCUM WHEN '' THEN 'TARGEN 73' ELSE 'TARG 73' END AS TARGEN,
           CARGO AS DETALLE,
           NOMBRE1 || ' ' || APELLIDO_PATERNO || ' ' || APELLIDO_MATERNO AS NOMBRE, MCA_INH,FEC_ACTU
    FROM TRON2000.TARGEN73
    WHERE MCA_INH <> 'S'

    UNION ALL

    -- lcfo_lst_negra
    SELECT TIPO_ID AS TIP_DOCUM, COD_ID AS COD_DOCUM, 
           CASE COD_ID WHEN '' THEN 'LCFO' ELSE 'LCFO LST NEGRA' END AS TARGEN,
           DETA_MOT AS DETALLE,
           NOM_COMPLETO AS NOMBRE,MCA_INH,FEC_ACTU
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
    
    odbcClose(c)
    
    resultado_cruce(cruce_listas)
    
    output$tabla_resultado <- renderDT({
      datatable(cruce_listas, options = list(pageLength = 10))
    })
  })

output$descargar_excel <- downloadHandler(
  filename = function() {
    paste("resultado_cruce_", Sys.Date(), ".xlsx", sep = "")
  },
  content = function(file) {
    write_xlsx(resultado_cruce(), path = file)
  }
)
}

shinyApp(ui = ui, server = server)

