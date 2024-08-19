################################################################################
############################SECCIÓN 1: LIBRERIAS################################
################################################################################
# Asegurarse de que las librerías necesarias están instaladas y cargadas
packages <- c("MASS", "lme4", "mice", "ggplot2", "gridExtra", "cluster", 
              "factoextra", "NbClust", "car", "caret", "sampling", 
              "purrr", "randomForest", "MuMIn", "dplyr", "survey", "corrplot", 
              "moments", "glmmTMB", "gbm", "naniar", "writexl", "png", "grid", 
              "e1071", "cowplot", "MuMIn", "pROC", "broom.mixed", "effects", "openxlsx", 
              "stringr", "AER", "lmtest", "pscl", "DHARMa", "tidyr", "sampling", "plotly", "ahpsurvey")

new.packages <- packages[!(packages %in% installed.packages()[,"Package"])]
if(length(new.packages)) install.packages(new.packages, dependencies = TRUE)

lapply(packages, library, character.only = TRUE)


################################################################################
##############SECCIÓN 2: PREPARACIÓN DE CONJUNTO DE DATOS#######################
################################################################################
# Configuración del directorio de trabajo y carga de datos
setwd("C:/Users/javer/OneDrive/Documentos/RESULTADOS TESIS/Resultados_Proyecto_FInal/")
datos <- read.csv("participacion.csv", stringsAsFactors = FALSE)
# Configurar el directorio de salida
directorio_salida <- "C:/Users/javer/OneDrive/Documentos/RESULTADOS TESIS/Resultados_Proyecto_FInal/"


establecimientos <- read.csv("establecimientos.csv", stringsAsFactors = FALSE)

# Unir datos con establecimientos usando la columna Cod_Vigente
datos <- merge(datos, establecimientos[, c("Cod_Vigente", "Cod_Comuna")], 
               by = "Cod_Vigente", all.x = TRUE)

# Actualizar la variable Mes para reflejar el mes continuo
datos$Ano <- as.numeric(datos$Ano)
datos$Mes <- as.numeric(datos$Mes)
datos <- datos %>%
  mutate(Mes = (Ano - 2021) * 12 + Mes)

# Normalizar las categorías de Tipo_Establec y preparar los datos
datos <- datos %>%                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 
  rename(Servicio_Salud = Dep_Jerarquica, Cod_Establecimiento = Cod_Vigente) %>%
  mutate(Tipo_Establec = str_trim(Tipo_Establec)) %>%
  mutate(Tipo_Establec = str_replace_all(Tipo_Establec, c(
    "Centro Comunitario de Salud Mental  \\(COSAM\\)" = "Centro Comunitario de Salud Mental",
    "Servicio de Atención Primaria de  Urgencia \\(SAPU\\)" = "Servicio de Atención Primaria de Urgencia (SAPU)",
    "Servicio de Atención Primaria de Urgencia de Alta Resolutividad  \\(SAR\\)" = "Servicio de Atención Primaria de Urgencia de Alta Resolutividad (SAR)"
  ))) %>%
  mutate(across(c(Cod_Establecimiento, Tipo_Establec, Cod_Region, Servicio_Salud, Nom_Comuna, 
                  Nivel_Atencion, Mes, Ano), as.factor)) %>%
  mutate(Nivel_Atencion = case_when(
    Nivel_Atencion %in% c("Hospital Comunitario con APS (Primario-Secundario)", 
                          "Hospital con APS (Primario y Secundario)", 
                          "Primario-Secundario ( Por confirmar)") ~ "Primario-Secundario",
    Nivel_Atencion %in% c("Secundario (Hospital Sin APS)", "Pendiente") ~ "Secundario",
    TRUE ~ as.character(Nivel_Atencion)
  )) %>%
  mutate(Nivel_Atencion = factor(Nivel_Atencion, levels = unique(Nivel_Atencion)))

# Verificar la recategorización
print(table(datos$Nivel_Atencion))

# Calcular la tasa de Participantes Por Actividad
datos <- datos %>%
  mutate(participantesxactividad = TotalParticipantes / TotalActividades,
         participantesxactividad = round(participantesxactividad))

# Identificar valores NaN
valores_nan <- is.na(datos$participantesxactividad)
datos$participantesxactividad[valores_nan & (datos$TotalParticipantes == 0 & datos$TotalActividades == 0)] <- 0
datos$participantesxactividad <- ifelse(datos$participantesxactividad > 0 & datos$participantesxactividad < 1, NA, datos$participantesxactividad)

# Identificar los Tipo_Establec con solo 0 en participantesxactividad
tipo_establec_solo_cero <- datos %>%
  group_by(Tipo_Establec) %>%
  summarise(todos_ceros = all(participantesxactividad == 0, na.rm = TRUE)) %>%
  filter(todos_ceros == TRUE)

# Mostrar los Tipo_Establec que tienen solo 0 en participantesxactividad
print(tipo_establec_solo_cero)

# Filtrar los datos para excluir los Tipo_Establec con solo 0 en participantesxactividad
datos_filtrados <- datos %>%
  filter(!(Tipo_Establec %in% tipo_establec_solo_cero$Tipo_Establec))

# Ajustar los niveles de los factores en el conjunto de datos filtrado
datos_filtrados <- datos_filtrados %>%
  mutate(across(where(is.factor), ~ factor(.x)))

# Calcular el porcentaje de observaciones con TotalParticipantes < TotalActividades
porcentaje_participantes_menor_actividades <- sum(datos$TotalParticipantes < datos$TotalActividades) / nrow(datos) * 100

# Calcular el porcentaje de observaciones con TotalOrganizaciones < TotalActividades
porcentaje_organizaciones_menor_actividades <- sum(datos$TotalOrganizaciones < datos$TotalActividades) / nrow(datos) * 100

# Mostrar los resultados
print(paste("Porcentaje de observaciones con participantes < actividades:", porcentaje_participantes_menor_actividades))
print(paste("Porcentaje de observaciones con organizaciones < actividades:", porcentaje_organizaciones_menor_actividades))

# Guardar el conjunto de datos filtrados en un archivo Excel
write.xlsx(datos_filtrados, file = paste0(directorio_salida, "datos_filtrados.xlsx"))

# Guardar un resumen de los datos en otro archivo Excel
resumen_datos <- summary(datos_filtrados)
write.xlsx(resumen_datos, file = paste0(directorio_salida, "resumen_datos.xlsx"))


################################################################################
################SECCIÓN 3: ANÁLISIS EXPLORATORIO DE DATOS#######################
################################################################################

# Crear una variable que indique si el establecimiento tiene al menos un registro mayor que 0
datos_filtrados <- datos_filtrados %>%
  mutate(sobre_cero = ifelse(TotalParticipantes > 0 | TotalActividades > 0 | TotalOrganizaciones > 0, "Mayor que 0", "0"))

sum(is.na(datos_filtrados$TotalParticipantes))
sum(is.na(datos_filtrados$TotalActividades))
sum(is.na(datos_filtrados$TotalOrganizaciones))
sum(is.na(datos_filtrados$Servicio_Salud))
sum(is.na(datos_filtrados$Tipo_Establec))
sum(is.na(datos_filtrados$Ano))
sum(is.na(datos_filtrados$Mes))

table(datos_filtrados$Servicio_Salud)
table(datos_filtrados$Tipo_Establec)
table(datos_filtrados$Ano)
table(datos_filtrados$Mes)

# Definir la ruta de salida
directorio_salida <- "C:/Users/javer/OneDrive/Documentos/RESULTADOS TESIS/Resultados_Proyecto_FInal/graficas_resultados_filtrados/"

# Crear el directorio si no existe
dir.create(directorio_salida, recursive = TRUE, showWarnings = FALSE)

# Gráfico de barras para variables categóricas con proporción en distinto color
categorical_vars <- c("Cod_Region", "Servicio_Salud", "Tipo_Establec", "Nivel_Atencion")

for (var in categorical_vars) {
  p_filtrados <- ggplot(datos_filtrados, aes(x = !!sym(var), fill = sobre_cero)) +
    geom_bar(position = "stack", color = "black") +
    scale_fill_manual(values = c("0" = "skyblue", "Mayor que 0" = "darkblue")) +
    labs(title = paste("Distribución de", var, "con Proporción sobre 0 (datos_filtrados)"),
         x = var,
         y = "Frecuencia") +
    theme_minimal() +
    theme(axis.text.x = element_text(angle = 45, hjust = 1))
  
  ggsave(filename = paste0(directorio_salida, "distribucion_proporcion_", var, ".png"), plot = p_filtrados, width = 8, height = 6)
}

# Cargar las imágenes de los gráficos y crear la matriz de gráficos
graficos_filtrados <- lapply(categorical_vars, function(var) {
  readPNG(paste0(directorio_salida, "distribucion_proporcion_", var, ".png"))
})

plots_filtrados <- lapply(graficos_filtrados, rasterGrob, interpolate = TRUE)
matriz_graficos_proporcion_filtrados <- do.call(grid.arrange, c(plots_filtrados, ncol = 2))

# Guardar la matriz de gráficos como una imagen
ggsave(paste0(directorio_salida, "matriz_graficos_proporcion.png"), plot = matriz_graficos_proporcion_filtrados, width = 16, height = 12)

# Lista de variables numéricas
numerical_vars <- c("TotalParticipantes", "TotalActividades", "TotalOrganizaciones")

# Crear listas para almacenar los gráficos
boxplot_list_filtrados <- list()

# Crear los diagramas de caja para cada variable numérica
for (var in numerical_vars) {
  p_filtrados <- ggplot(datos_filtrados, aes(x = "1", y = .data[[var]])) +
    geom_boxplot(fill = "lightblue", color = "black") +
    labs(title = paste("Diagrama de Caja de", var, "(datos_filtrados)"),
         x = "",
         y = var) +
    theme_minimal()
  
  boxplot_list_filtrados[[var]] <- p_filtrados
}

# Crear la matriz de gráficos
matriz_boxplot_filtrados <- do.call(grid.arrange, c(boxplot_list_filtrados, ncol = 2))

# Guardar la matriz de gráficos como una imagen
ggsave(paste0(directorio_salida, "matriz_boxplot.png"), plot = matriz_boxplot_filtrados, width = 16, height = 12)

# Gráfico de correlación
# Calcular la matriz de correlación para datos_filtrados
cor_matrix_filtrados <- cor(datos_filtrados[, numerical_vars], use = "complete.obs")

# Generar gráficos de correlación y guardarlos como imágenes
png(filename = paste0(directorio_salida, "correlacion.png"), width = 800, height = 600)
corrplot(cor_matrix_filtrados, method = "circle", type = "lower", tl.cex = 0.8)
dev.off()

# Gráfico de evolución de Variables Numéricas
grafico_evolucion <- datos_filtrados %>%
  group_by(Ano, Mes) %>%
  summarise(
    TotalParticipantes = sum(TotalParticipantes),
    TotalActividades = sum(TotalActividades),
    TotalOrganizaciones = sum(TotalOrganizaciones),
    .groups = 'drop' # Para evitar el warning sobre el agrupamiento
  ) %>%
  pivot_longer(cols = c(TotalParticipantes, TotalActividades, TotalOrganizaciones), 
               names_to = "Variable", values_to = "Valor") %>%
  ggplot(aes(x = interaction(Ano, Mes, sep = "-"), y = Valor, color = Variable, group = Variable)) +
  geom_line() +
  geom_point() +
  labs(title = "Evolución de Variables Numéricas",
       x = "Tiempo (Año-Mes)",
       y = "Valor",
       color = "Variable") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

# Guardar el gráfico
ggsave(filename = paste0(directorio_salida, "evolucion_variables_numericas.png"), plot = grafico_evolucion, width = 10, height = 6)



## Evaluación de sobredispercion
# Calcular media y varianza
mean_val <- mean(datos_filtrados$TotalOrganizaciones)
var_val <- var(datos_filtrados$TotalOrganizaciones)
print(c("Media" = mean_val, "Varianza" = var_val))



################################################################################
################################################################################
###############################MUESTREO#########################################
################################################################################
################################################################################


### SELECCIÓN MESES PARA TotalOrganizaciones ###

# Calcular la varianza de las observaciones por mes para TotalOrganizaciones
varianza_organizaciones_por_mes <- datos_filtrados %>%
  group_by(Mes) %>%
  summarise(Varianza = var(TotalOrganizaciones))

# Identificar los meses clave para TotalOrganizaciones

# Mes con mayor varianza
mes_mayor_varianza_organizaciones <- varianza_organizaciones_por_mes %>%
  filter(Varianza == max(Varianza)) %>%
  pull(Mes)

# Mes más cercano al promedio de la variabilidad
promedio_varianza_organizaciones <- mean(varianza_organizaciones_por_mes$Varianza)
mes_mas_cercano_promedio_organizaciones <- varianza_organizaciones_por_mes %>%
  mutate(Distancia_Al_Promedio = abs(Varianza - promedio_varianza_organizaciones)) %>%
  filter(Distancia_Al_Promedio == min(Distancia_Al_Promedio)) %>%
  pull(Mes)

# Mostrar resultados
cat("TotalOrganizaciones:\n")
cat("Mes con mayor variabilidad:", mes_mayor_varianza_organizaciones, "\n")
cat("Mes con variabilidad más cercana al promedio:", mes_mas_cercano_promedio_organizaciones, "\n\n")

### SELECCIÓN DE SUBCONJUNTOS ###

# Función para seleccionar un subconjunto del mes especificado
seleccionar_mes_subset <- function(datos, mes_seleccionado) {
  datos %>%
    filter(Mes == mes_seleccionado) %>%
    group_by(Cod_Establecimiento) %>%
    slice(1) %>%
    ungroup()
}

# Crear subconjuntos para los meses seleccionados para TotalOrganizaciones
maxvar_organizaciones <- seleccionar_mes_subset(datos_filtrados, mes_mayor_varianza_organizaciones)
promvar_organizaciones <- seleccionar_mes_subset(datos_filtrados, mes_mas_cercano_promedio_organizaciones)

# Contar el número de establecimientos en cada subconjunto
n_establecimientos_maxvar <- n_distinct(maxvar_organizaciones$Cod_Establecimiento)
n_establecimientos_promvar <- n_distinct(promvar_organizaciones$Cod_Establecimiento)

# Mostrar los resultados
cat("Número de establecimientos en el mes con mayor variabilidad:", n_establecimientos_maxvar, "\n")
cat("Número de establecimientos en el mes con variabilidad más cercana al promedio:", n_establecimientos_promvar, "\n\n")

# Mostrar tablas de frecuencias para otras variables
table(promvar_organizaciones$Servicio_Salud)
table(promvar_organizaciones$Cod_Region)
table(promvar_organizaciones$Tipo_Establec)
table(promvar_organizaciones$Nivel_Atencion)

table(maxvar_organizaciones$Servicio_Salud)
table(maxvar_organizaciones$Cod_Region)
table(maxvar_organizaciones$Tipo_Establec)
table(maxvar_organizaciones$Nivel_Atencion)

### Obtener y mostrar resúmenes ###

# TotalOrganizaciones - Mes con mayor variabilidad
resumen_maxvar_organizaciones <- summary(maxvar_organizaciones$TotalOrganizaciones)
var_maxvar_organizaciones <- var(maxvar_organizaciones$TotalOrganizaciones)
sd_maxvar_organizaciones <- sd(maxvar_organizaciones$TotalOrganizaciones)

cat("Resumen de TotalOrganizaciones para el mes con mayor variabilidad (", mes_mayor_varianza_organizaciones, "):\n", resumen_maxvar_organizaciones, "\n")
cat("Varianza de TotalOrganizaciones para el mes con mayor variabilidad:", var_maxvar_organizaciones, "\n")
cat("Desviación estándar de TotalOrganizaciones para el mes con mayor variabilidad:", sd_maxvar_organizaciones, "\n\n")

# TotalOrganizaciones - Mes más cercano al promedio
resumen_promvar_organizaciones <- summary(promvar_organizaciones$TotalOrganizaciones)
var_promvar_organizaciones <- var(promvar_organizaciones$TotalOrganizaciones)
sd_promvar_organizaciones <- sd(promvar_organizaciones$TotalOrganizaciones)

cat("Resumen de TotalOrganizaciones para el mes más cercano al promedio (", mes_mas_cercano_promedio_organizaciones, "):\n", resumen_promvar_organizaciones, "\n")
cat("Varianza de TotalOrganizaciones para el mes más cercano al promedio:", var_promvar_organizaciones, "\n")
cat("Desviación estándar de TotalOrganizaciones para el mes más cercano al promedio:", sd_promvar_organizaciones, "\n\n")


###############################ALEATORIO########################################
################################################################################
# Inicializar listas para almacenar resultados y simulaciones
resultados_estadisticas <- list()
resultados_simulaciones <- list()

# 1. Funciones Utilitarias

# 1.1. Función para realizar el muestreo aleatorio simple sin reemplazo
realizar_muestreo_aleatorio_simple <- function(datos, tamano_muestra, variable_interes) {
  muestra <- datos %>%
    slice_sample(n = tamano_muestra)
  
  mean(muestra[[variable_interes]])
}

# 1.2. Función para calcular estadísticas, incluyendo el cálculo de la varianza
calcular_estadisticas <- function(resultados, n, valor_verdadero, metodo) {
  media_estimaciones <- mean(resultados)
  sesgo <- media_estimaciones - valor_verdadero
  varianza <- var(resultados)
  error_cuadratico_medio <- mean((resultados - valor_verdadero)^2)
  
  data.frame(
    Metodo = metodo,
    Tamano_Muestra = n,
    Media_Estimaciones = media_estimaciones,
    Sesgo = sesgo,
    Varianza = varianza,
    ECM = error_cuadratico_medio,
    Valor_Real = valor_verdadero
  )
}
# Subconjuntos para TotalOrganizaciones
meses_organizaciones <- list(
  maxvar_organizaciones = maxvar_organizaciones,
  promvar_organizaciones = promvar_organizaciones
)

# Procesar TotalOrganizaciones
for (subset_name in names(meses_organizaciones)) {
  mes_subset <- meses_organizaciones[[subset_name]]
  variable <- "TotalOrganizaciones"
  valor_verdadero <- mean(mes_subset[[variable]])
  mes <- unique(mes_subset$Mes)
  
  set.seed(123)
  resultados_100 <- replicate(1000, realizar_muestreo_aleatorio_simple(mes_subset, 100, variable))
  resultados_200 <- replicate(1000, realizar_muestreo_aleatorio_simple(mes_subset, 200, variable))
  
  metodo <- "Aleatorio_Simple"
  
  # Calcular estadísticas
  estadisticas_100 <- calcular_estadisticas(resultados_100, 100, valor_verdadero, metodo) %>%
    mutate(Mes = mes, Subset = subset_name)
  estadisticas_200 <- calcular_estadisticas(resultados_200, 200, valor_verdadero, metodo) %>%
    mutate(Mes = mes, Subset = subset_name)
  
  resultados_estadisticas[[paste0(variable, "_", subset_name, "_100")]] <- estadisticas_100
  resultados_estadisticas[[paste0(variable, "_", subset_name, "_200")]] <- estadisticas_200
  
  # Agregar una columna para identificar el método y consolidar simulaciones
  simulaciones_100 <- data.frame(Valor_Simulado = resultados_100) %>%
    mutate(Tamano_Muestra = 100, Variable = variable, Mes = mes, Metodo = metodo, Subset = subset_name, Valor_Real = valor_verdadero)
  
  simulaciones_200 <- data.frame(Valor_Simulado = resultados_200) %>%
    mutate(Tamano_Muestra = 200, Variable = variable, Mes = mes, Metodo = metodo, Subset = subset_name, Valor_Real = valor_verdadero)
  
  resultados_simulaciones[[paste0(variable, "_", subset_name, "_100")]] <- simulaciones_100
  resultados_simulaciones[[paste0(variable, "_", subset_name, "_200")]] <- simulaciones_200
}
# 3.1 Consolidar los resultados en un solo dataframe
todas_estadisticas <- do.call(rbind, resultados_estadisticas)

# 3.2 Consolidar las simulaciones en un solo dataframe
todas_simulaciones <- do.call(rbind, resultados_simulaciones)

# 3.3 Crear un archivo Excel con los resultados combinados
wb <- createWorkbook()

addWorksheet(wb, "Resultados_Combinados")
writeData(wb, "Resultados_Combinados", todas_estadisticas)

addWorksheet(wb, "Simulaciones_Combinadas")
writeData(wb, "Simulaciones_Combinadas", todas_simulaciones)

saveWorkbook(wb, file.path(directorio_salida, "Resultados_aleatorio.xlsx"), overwrite = TRUE)

print("Resultados consolidados guardados exitosamente en el directorio especificado.")
# 4. Crear una función para generar gráficos de histogramas y organizarlos en una matriz gráfica
crear_matriz_graficos <- function(simulaciones, metodo) {
  graficos <- list()  # Lista para almacenar los gráficos
  
  for (subset_name in unique(simulaciones$Subset)) {
    for (tamano_muestra in unique(simulaciones$Tamano_Muestra)) {
      datos_filtrados <- simulaciones %>%
        filter(Subset == subset_name, Tamano_Muestra == tamano_muestra)
      
      # Obtener el valor real del parámetro
      valor_real <- unique(datos_filtrados$Valor_Real)[1]
      
      if (length(valor_real) == 0 || is.na(valor_real)) {
        next  # Si no se encuentra el valor real, continuar con el siguiente gráfico
      }
      
      grafico <- ggplot(datos_filtrados, aes(x = Valor_Simulado)) +
        geom_histogram(binwidth = 0.5, fill = "blue", color = "black", alpha = 0.7) +
        geom_vline(xintercept = valor_real, color = "red", linetype = "dashed", size = 1) +  # Añadir la línea vertical
        labs(
          title = paste("Distribución de Simulaciones"),
          subtitle = paste("Mes:", unique(datos_filtrados$Mes), "- Subset:", subset_name, "- Tamaño de Muestra:", tamano_muestra),
          x = "Valor Simulado",
          y = "Frecuencia"
        ) +
        theme_minimal()
      
      graficos[[paste0(subset_name, "_", tamano_muestra)]] <- grafico
    }
  }
  
  # Crear una matriz gráfica con todos los gráficos generados
  matriz_graficos <- marrangeGrob(graficos, nrow = 2, ncol = 2)  # Ajusta nrow y ncol según la cantidad de gráficos y el tamaño deseado
  
  return(matriz_graficos)
}

# 5. Llamar a la función para generar la matriz de gráficos
matriz_graficos <- crear_matriz_graficos(todas_simulaciones, "Aleatorio_Simple")

# 6. Guardar la matriz de gráficos como un archivo de imagen
nombre_archivo <- file.path(directorio_salida, "Matriz_Graficos_Aleatorio_Simple.png")
ggsave(filename = nombre_archivo, plot = matriz_graficos, width = 14, height = 10, dpi = 300)

print("Matriz de gráficos generada y guardada exitosamente en el directorio especificado.")



###############################ESTRATIFICADO####################################
################################################################################
################################################################################
# 1.1. Función para asignación proporcional
asignacion_proporcional <- function(datos, n, estrato) {
  tabla_estratos <- table(datos[[estrato]])
  proporciones <- prop.table(tabla_estratos)
  tamanos_muestra <- round(proporciones * n)
  return(tamanos_muestra)
}

# 1.2. Función para asignación de Neyman
asignacion_neyman <- function(datos, n, estrato, variable_interes) {
  tabla_estratos <- table(datos[[estrato]])
  N_h <- as.numeric(tabla_estratos)
  sd_h <- tapply(datos[[variable_interes]], datos[[estrato]], sd, na.rm = TRUE)
  numerador <- N_h * sd_h
  sum_numerador <- sum(numerador, na.rm = TRUE)
  tamanos_muestra <- round((numerador / sum_numerador) * n)
  return(tamanos_muestra)
}

# 1.3. Función para realizar el muestreo estratificado
realizar_muestreo_estratificado <- function(datos, n, metodo, estrato, variable_interes) {
  metodo <- tolower(metodo)  # Convertir el método a minúsculas para evitar problemas de comparación
  if (metodo == "proporcional") {
    tamanos_muestra <- asignacion_proporcional(datos, n, estrato)
  } else if (metodo == "neyman") {
    tamanos_muestra <- asignacion_neyman(datos, n, estrato, variable_interes)
  } else {
    stop("Método de asignación no reconocido. Use 'Proporcional' o 'Neyman'.")
  }
  
  niveles_estrato <- names(tamanos_muestra)
  muestra <- do.call(rbind, lapply(niveles_estrato, function(nivel) {
    datos %>%
      filter(datos[[estrato]] == nivel) %>%
      slice_sample(n = tamanos_muestra[nivel])
  }))
  
  return(mean(muestra[[variable_interes]]))
}

# 1.4. Función para calcular estadísticas incluyendo el cálculo del efecto de diseño
calcular_estadisticas_estratificado <- function(resultados, n, valor_verdadero, varianza_simple, estrato, metodo, mes, variable) {
  media_estimaciones <- mean(resultados)
  sesgo <- media_estimaciones - valor_verdadero
  varianza <- var(resultados)
  error_cuadratico_medio <- mean((resultados - valor_verdadero)^2)
  
  # Calcular efecto de diseño
  efecto_diseno <- varianza / varianza_simple
  
  data.frame(
    Metodo = metodo,
    Estrato = estrato,
    Mes = mes,
    Variable = variable,
    Tamano_Muestra = n,
    Media_Estimaciones = media_estimaciones,
    Sesgo = sesgo,
    Varianza = varianza,
    ECM = error_cuadratico_medio,
    Efecto_Diseno = efecto_diseno
  )
}

# 2.1. Inicializar listas para almacenar resultados y simulaciones
resultados_estadisticas_estratificado <- list()
resultados_simulaciones_estratificado <- list()

# 2.2. Definir la variable de interés y variables de estratificación
variable_interes <- "TotalOrganizaciones"
variables_estratificacion <- list("Nivel_Atencion", "Tipo_Establec", "Servicio_Salud", "Cod_Region")

# 2.3. Definir subconjuntos de meses específicos para TotalOrganizaciones
meses_subsets <- list(
  maxvar = maxvar_organizaciones,
  promvar = promvar_organizaciones
)

# 2.4. Realizar simulaciones y calcular estadísticas
for (estrato in variables_estratificacion) {
  for (subset_name in names(meses_subsets)) {
    mes_subset <- meses_subsets[[subset_name]]
    valor_verdadero <- mean(mes_subset[[variable_interes]])
    varianza_simple <- var(replicate(1000, realizar_muestreo_estratificado(mes_subset, 100, "Proporcional", estrato, variable_interes)))
    mes <- unique(mes_subset$Mes)
    
    # Realizar 1000 simulaciones para cada tamaño de muestra y método
    set.seed(123)
    resultados_100_prop <- replicate(1000, realizar_muestreo_estratificado(mes_subset, 100, "Proporcional", estrato, variable_interes))
    resultados_200_prop <- replicate(1000, realizar_muestreo_estratificado(mes_subset, 200, "Proporcional", estrato, variable_interes))
    resultados_100_neyman <- replicate(1000, realizar_muestreo_estratificado(mes_subset, 100, "Neyman", estrato, variable_interes))
    resultados_200_neyman <- replicate(1000, realizar_muestreo_estratificado(mes_subset, 200, "Neyman", estrato, variable_interes))
    
    # Calcular estadísticas para ambos tamaños de muestra y métodos
    estadisticas_100_prop <- calcular_estadisticas_estratificado(resultados_100_prop, 100, valor_verdadero, varianza_simple, estrato, "Proporcional", mes, variable_interes)
    estadisticas_200_prop <- calcular_estadisticas_estratificado(resultados_200_prop, 200, valor_verdadero, varianza_simple, estrato, "Proporcional", mes, variable_interes)
    estadisticas_100_neyman <- calcular_estadisticas_estratificado(resultados_100_neyman, 100, valor_verdadero, varianza_simple, estrato, "Neyman", mes, variable_interes)
    estadisticas_200_neyman <- calcular_estadisticas_estratificado(resultados_200_neyman, 200, valor_verdadero, varianza_simple, estrato, "Neyman", mes, variable_interes)
    
    # Almacenar resultados en listas
    resultados_estadisticas_estratificado[[paste0(variable_interes, "_P_100_M", mes, "_", estrato, "_", subset_name)]] <- estadisticas_100_prop
    resultados_estadisticas_estratificado[[paste0(variable_interes, "_P_200_M", mes, "_", estrato, "_", subset_name)]] <- estadisticas_200_prop
    resultados_estadisticas_estratificado[[paste0(variable_interes, "_N_100_M", mes, "_", estrato, "_", subset_name)]] <- estadisticas_100_neyman
    resultados_estadisticas_estratificado[[paste0(variable_interes, "_N_200_M", mes, "_", estrato, "_", subset_name)]] <- estadisticas_200_neyman
    
    # Almacenar simulaciones en listas con identificación
    simulaciones_100_prop <- data.frame(Valor_Simulado = resultados_100_prop) %>%
      mutate(Tamano_Muestra = 100, Variable = variable_interes, Mes = mes, Estrato = estrato, Metodo = "Proporcional", Subset = subset_name)
    
    simulaciones_200_prop <- data.frame(Valor_Simulado = resultados_200_prop) %>%
      mutate(Tamano_Muestra = 200, Variable = variable_interes, Mes = mes, Estrato = estrato, Metodo = "Proporcional", Subset = subset_name)
    
    simulaciones_100_neyman <- data.frame(Valor_Simulado = resultados_100_neyman) %>%
      mutate(Tamano_Muestra = 100, Variable = variable_interes, Mes = mes, Estrato = estrato, Metodo = "Neyman", Subset = subset_name)
    
    simulaciones_200_neyman <- data.frame(Valor_Simulado = resultados_200_neyman) %>%
      mutate(Tamano_Muestra = 200, Variable = variable_interes, Mes = mes, Estrato = estrato, Metodo = "Neyman", Subset = subset_name)
    
    resultados_simulaciones_estratificado[[paste0(variable_interes, "_P_100_M", mes, "_", estrato, "_", subset_name)]] <- simulaciones_100_prop
    resultados_simulaciones_estratificado[[paste0(variable_interes, "_P_200_M", mes, "_", estrato, "_", subset_name)]] <- simulaciones_200_prop
    resultados_simulaciones_estratificado[[paste0(variable_interes, "_N_100_M", mes, "_", estrato, "_", subset_name)]] <- simulaciones_100_neyman
    resultados_simulaciones_estratificado[[paste0(variable_interes, "_N_200_M", mes, "_", estrato, "_", subset_name)]] <- simulaciones_200_neyman
  }
}

# 3.1. Consolidar todos los resultados en un solo data.frame
estadisticas_completas <- do.call(rbind, resultados_estadisticas_estratificado)

# 3.2. Consolidar todas las simulaciones en un solo data.frame
simulaciones_completas <- do.call(rbind, resultados_simulaciones_estratificado)

# 3.3. Crear un archivo Excel con los resultados combinados
wb <- createWorkbook()

# Crear y escribir la hoja para los resultados de estadísticas
addWorksheet(wb, "Resultados_Estratificado")
writeData(wb, "Resultados_Estratificado", estadisticas_completas)

# Crear y escribir la hoja para los resultados de simulaciones
addWorksheet(wb, "Simulaciones_Estratificado")
writeData(wb, "Simulaciones_Estratificado", simulaciones_completas)

# Guardar el archivo Excel en el directorio de salida
saveWorkbook(wb, file.path(directorio_salida, "Resultados_Estratificado.xlsx"), overwrite = TRUE)

# Mensaje de confirmación
print("Resultados guardados exitosamente en el directorio especificado.")

# 4.1. Crear una función para generar gráficos y guardarlos en el directorio
crear_graficos_simulaciones <- function(simulaciones_completas, directorio_salida, valor_verdadero) {
  # Iterar sobre cada variable de estratificación
  for (estrato in unique(simulaciones_completas$Estrato)) {
    
    # Filtrar las simulaciones para el estrato actual
    simulaciones_estrato <- simulaciones_completas %>% filter(Estrato == estrato)
    
    # Crear los gráficos para cada combinación de tamaño de muestra, método y subset
    graficos <- list()
    
    for (tamano_muestra in c(100, 200)) {
      for (metodo in c("Proporcional", "Neyman")) {
        for (subset_name in c("promvar", "maxvar")) {
          
          simulaciones_subset <- simulaciones_estrato %>%
            filter(Tamano_Muestra == tamano_muestra, Metodo == metodo, Subset == subset_name)
          
          # Crear el gráfico con un binwidth más pequeño para barras más finas
          p <- ggplot(simulaciones_subset, aes(x = Valor_Simulado)) +
            geom_histogram(binwidth = 0.5, fill = "blue", alpha = 0.7) +
            geom_vline(aes(xintercept = valor_verdadero), color = "red", linetype = "dashed", size = 1) +
            labs(
              title = paste0("Muestra: ", tamano_muestra, ", Método: ", metodo, ", Subset: ", subset_name),
              x = "Valor Simulado",
              y = "Frecuencia"
            ) +
            theme_minimal()
          
          graficos[[paste0(metodo, "_", tamano_muestra, "_", subset_name)]] <- p
        }
      }
    }
    
    # Crear una matriz de gráficos (2x4)
    matriz_graficos <- grid.arrange(
      graficos$Proporcional_100_promvar, graficos$Neyman_100_promvar,
      graficos$Proporcional_100_maxvar, graficos$Neyman_100_maxvar,
      graficos$Proporcional_200_promvar, graficos$Neyman_200_promvar,
      graficos$Proporcional_200_maxvar, graficos$Neyman_200_maxvar,
      ncol = 2, nrow = 4,
      top = paste0("Distribución de Simulaciones para el Estrato: ", estrato)
    )
    
    # Guardar la matriz de gráficos en un archivo de imagen
    nombre_archivo <- file.path(directorio_salida, paste0("Graficos_Simulaciones_", estrato, ".png"))
    ggsave(nombre_archivo, matriz_graficos, width = 16, height = 20)
  }
  
  # Mensaje de confirmación
  print("Gráficos guardados exitosamente en el directorio especificado.")
}

# 4.2. Ejecutar la función para crear y guardar los gráficos
crear_graficos_simulaciones(simulaciones_completas, directorio_salida, valor_verdadero)


###########################MUESTREOS CONGLOMERADOS#############################
################################################################################
################################################################################


# Inicializar listas para almacenar resultados y simulaciones
resultados_estadisticas_conglomerados <- list()
resultados_simulaciones_conglomerados <- list()

# 1. Función para realizar el muestreo por conglomerados
realizar_muestreo_conglomerados <- function(datos, n, variable_interes) {
  comunas <- unique(datos$Nom_Comuna)
  
  # Inicializar variables
  muestra_total <- NULL
  total_seleccionado <- 0
  n_conglomerados <- 1
  
  # Iterar hasta que se alcance el tamaño de muestra deseado o se hayan seleccionado todas las comunas
  while (total_seleccionado < n && n_conglomerados <= length(comunas)) {
    # Seleccionar aleatoriamente un número de comunas
    comunas_seleccionadas <- sample(comunas, n_conglomerados, replace = FALSE)
    
    # Obtener todos los establecimientos dentro de las comunas seleccionadas
    muestra <- datos %>% filter(Nom_Comuna %in% comunas_seleccionadas)
    
    # Calcular el tamaño de la muestra seleccionada
    total_seleccionado <- nrow(muestra)
    
    # Guardar la muestra seleccionada
    muestra_total <- muestra
    
    # Incrementar el número de conglomerados
    n_conglomerados <- n_conglomerados + 1
  }
  
  # Calcular la media de la variable de interés en la muestra
  mean(muestra_total[[variable_interes]])
}

# 2. Función para calcular estadísticas, incluyendo el cálculo del efecto de diseño
calcular_estadisticas_conglomerados <- function(resultados, n, conglomerado, variable, valor_verdadero, mes, varianza_simple) {
  media_estimaciones <- mean(resultados)
  sesgo <- media_estimaciones - valor_verdadero
  varianza <- var(resultados)
  error_cuadratico_medio <- mean((resultados - valor_verdadero)^2)
  
  # Calcular el efecto de diseño
  efecto_diseno <- varianza / varianza_simple
  
  data.frame(
    Conglomerado = conglomerado,
    Mes = mes,
    Variable = variable,
    Tamano_Muestra = n,
    Media_Estimaciones = media_estimaciones,
    Sesgo = sesgo,
    Varianza = varianza,
    ECM = error_cuadratico_medio,
    Efecto_Diseno = efecto_diseno
  )
}

# 3. Proceso de Muestreo y Cálculo de Estadísticas

# 3.1. Definir la variable de interés
variable_interes <- "TotalOrganizaciones"

# Subconjuntos de meses para TotalOrganizaciones
meses_subsets <- list(
  maxvar_organizaciones = maxvar_organizaciones,
  promvar_organizaciones = promvar_organizaciones
)

for (subset_name in names(meses_subsets)) {
  mes_subset <- meses_subsets[[subset_name]]
  valor_verdadero <- mean(mes_subset[[variable_interes]])
  mes <- unique(mes_subset$Mes)
  
  # Calcular varianza de muestreo aleatorio simple como referencia para el efecto de diseño
  varianza_simple <- var(replicate(1000, realizar_muestreo_aleatorio_simple(mes_subset, 100, variable_interes)))
  
  # Realizar 1000 simulaciones para cada tamaño de muestra
  set.seed(123)  # Para reproducibilidad
  resultados_100 <- replicate(1000, realizar_muestreo_conglomerados(mes_subset, 100, variable_interes))
  resultados_200 <- replicate(1000, realizar_muestreo_conglomerados(mes_subset, 200, variable_interes))
  
  # Calcular estadísticas para ambos tamaños de muestra
  estadisticas_100 <- calcular_estadisticas_conglomerados(resultados_100, 100, "Nom_Comuna", variable_interes, valor_verdadero, mes, varianza_simple)
  estadisticas_200 <- calcular_estadisticas_conglomerados(resultados_200, 200, "Nom_Comuna", variable_interes, valor_verdadero, mes, varianza_simple)
  
  # Almacenar los resultados
  resultados_estadisticas_conglomerados[[paste0(variable_interes, "_Cong_", mes, "_100")]] <- estadisticas_100
  resultados_estadisticas_conglomerados[[paste0(variable_interes, "_Cong_", mes, "_200")]] <- estadisticas_200
  
  # Almacenar simulaciones con identificación
  simulaciones_100 <- data.frame(Valor_Simulado = resultados_100) %>%
    mutate(Tamano_Muestra = 100, Variable = variable_interes, Mes = mes, Conglomerado = "Nom_Comuna")
  
  simulaciones_200 <- data.frame(Valor_Simulado = resultados_200) %>%
    mutate(Tamano_Muestra = 200, Variable = variable_interes, Mes = mes, Conglomerado = "Nom_Comuna")
  
  resultados_simulaciones_conglomerados[[paste0(variable_interes, "_Cong_", mes, "_100")]] <- simulaciones_100
  resultados_simulaciones_conglomerados[[paste0(variable_interes, "_Cong_", mes, "_200")]] <- simulaciones_200
}

# 4. Generación de Archivos de Salida

# 4.1. Consolidar todos los resultados en un solo data.frame
estadisticas_completas <- do.call(rbind, resultados_estadisticas_conglomerados)

# 4.2. Consolidar todas las simulaciones en un solo data.frame
simulaciones_completas <- do.call(rbind, resultados_simulaciones_conglomerados)

# 4.3. Crear un archivo Excel con los resultados combinados
wb <- createWorkbook()

addWorksheet(wb, "Resultados_Conglomerados")
writeData(wb, "Resultados_Conglomerados", estadisticas_completas)

addWorksheet(wb, "Simulaciones_Conglomerados")
writeData(wb, "Simulaciones_Conglomerados", simulaciones_completas)

saveWorkbook(wb, file.path(directorio_salida, "Resultados_Conglomerados.xlsx"), overwrite = TRUE)

print("Resultados guardados exitosamente en el directorio especificado.")

# Función para crear gráficos de las distribuciones de las estimaciones
crear_graficos_distribuciones_conglomerados <- function(simulaciones, valor_real, binwidth = 0.5) {
  ggplot(simulaciones, aes(x = Valor_Simulado)) +
    geom_histogram(binwidth = binwidth, fill = "green", color = "black", alpha = 0.7) +
    geom_vline(xintercept = valor_real, color = "red", linetype = "dashed", size = 1) +
    facet_grid(Mes ~ Tamano_Muestra, scales = "free") +
    labs(
      title = "Distribución de Estimaciones por Conglomerado",
      x = "Valor Simulado",
      y = "Frecuencia"
    ) +
    theme_minimal()
}

# Generar gráficos para cada subset con un binwidth más pequeño
graficos <- list()

for (subset_name in names(meses_subsets)) {
  simulaciones_subset <- simulaciones_completas %>% filter(Mes %in% meses_subsets[[subset_name]]$Mes)
  valor_real <- mean(simulaciones_subset$Valor_Simulado)  # Aquí asumimos que el valor real es la media de las simulaciones
  
  grafico <- crear_graficos_distribuciones_conglomerados(simulaciones_subset, valor_real, binwidth = 0.1)
  graficos[[subset_name]] <- grafico
  
  # Guardar el gráfico
  ggsave(filename = file.path(directorio_salida, paste0("Grafico_Distribucion_", subset_name, ".png")), plot = grafico)
}

# Mostrar gráficos en el entorno de trabajo (si se desea visualizar inmediatamente)
for (grafico in graficos) {
  print(grafico)
}

print("Gráficos de distribuciones para conglomerados generados y guardados exitosamente.")


################################################################################
#################################multietapico2cong##############################
################################################################################

# 1. Funciones

## 1.1 Función para realizar el muestreo multietápico por comuna dentro de un estrato
muestreomulti2cong <- function(datos, n_estrato, variable_interes) {
  estratos <- unique(datos$estrato)
  muestra_total <- NULL
  
  for (estrato in estratos) {
    datos_estrato <- datos %>% filter(estrato == !!estrato)
    tamanos_muestra_estrato <- round((n_estrato * nrow(datos_estrato)) / nrow(datos))
    
    comunas <- unique(datos_estrato$Nom_Comuna)
    muestra_estrato <- NULL
    total_seleccionado <- 0
    
    while (total_seleccionado < tamanos_muestra_estrato && length(comunas) > 0) {
      comuna_seleccionada <- sample(comunas, 1)
      comunas <- setdiff(comunas, comuna_seleccionada)
      
      muestra_conglomerado <- datos_estrato %>% filter(Nom_Comuna == comuna_seleccionada)
      total_seleccionado <- total_seleccionado + nrow(muestra_conglomerado)
      
      muestra_estrato <- rbind(muestra_estrato, muestra_conglomerado)
    }
    
    if (total_seleccionado > tamanos_muestra_estrato) {
      exceso <- total_seleccionado - tamanos_muestra_estrato
      muestra_estrato <- muestra_estrato[-(1:exceso), ]
    }
    
    muestra_total <- rbind(muestra_total, muestra_estrato)
  }
  
  mean(muestra_total[[variable_interes]])
}

## 1.2 Función para calcular estadísticas, incluyendo el efecto de diseño
calculo_estadisticos_multi2cong <- function(resultados, n, valor_verdadero, estrato, mes, varianza_simple, conglomerado, variable) {
  media_estimaciones <- mean(resultados)
  sesgo <- media_estimaciones - valor_verdadero
  varianza <- var(resultados)
  error_cuadratico_medio <- mean((resultados - valor_verdadero)^2)
  
  # Calcular efecto de diseño
  efecto_diseno <- varianza / varianza_simple
  
  data.frame(
    Estrato = estrato,
    Mes = mes,
    Variable = variable,
    Tamano_Muestra = n,
    Media_Estimaciones = media_estimaciones,
    Sesgo = sesgo,
    Varianza = varianza,
    ECM = error_cuadratico_medio,
    Efecto_Diseno = efecto_diseno,
    Conglomerado = conglomerado
  )
}

## 1.3 Función para procesar el muestreo multietápico y calcular estadísticas
procesar_muestreo_multi2cong <- function(mes_subset, estrato, variable, varianza_simple, conglomerado) {
  mes_subset <- mes_subset %>% mutate(estrato = get(estrato))
  valor_verdadero <- mean(mes_subset[[variable]])
  mes <- unique(mes_subset$Mes)
  
  resultados_100 <- replicate(1000, muestreomulti2cong(mes_subset, 100, variable))
  resultados_200 <- replicate(1000, muestreomulti2cong(mes_subset, 200, variable))
  
  estadisticas_100 <- calculo_estadisticos_multi2cong(resultados_100, 100, valor_verdadero, estrato, mes, varianza_simple, conglomerado, variable)
  estadisticas_200 <- calculo_estadisticos_multi2cong(resultados_200, 200, valor_verdadero, estrato, mes, varianza_simple, conglomerado, variable)
  
  list(
    estadisticas_100 = estadisticas_100,
    estadisticas_200 = estadisticas_200,
    resultados_100 = resultados_100,
    resultados_200 = resultados_200
  )
}

# 2. Configuración

## 2.1 Definición de estratos y variable de interés
estratos <- c("Nivel_Atencion", "Tipo_Establec", "Servicio_Salud", "Cod_Region")
variable_interes <- "TotalOrganizaciones"
meses_subsets <- list(
  maxvar_organizaciones = maxvar_organizaciones,
  promvar_organizaciones = promvar_organizaciones
)

## 2.2 Inicialización de listas para almacenar resultados
resultados_estadisticos_multi2cong <- list()
resultados_simulaciones_multi2cong <- list()

## 2.3 Obtención de la varianza del muestreo aleatorio simple para cada subconjunto de mes
varianza_simple_list <- sapply(meses_subsets, function(mes_subset) {
  var(replicate(1000, realizar_muestreo_aleatorio_simple(mes_subset, 100, variable_interes)))
})

# 3. Proceso de Muestreo y Cálculo de Estadísticas

for (subset_name in names(meses_subsets)) {
  mes_subset <- meses_subsets[[subset_name]]
  for (estrato in estratos) {
    varianza_simple <- varianza_simple_list[subset_name]
    conglomerado <- "Nom_Comuna"
    resultado <- procesar_muestreo_multi2cong(mes_subset, estrato, variable_interes, varianza_simple, conglomerado)
    
    # Combinar resultados en data.frames únicos
    resultados_estadisticos_multi2cong[[paste0(variable_interes, "_", mes_subset$Mes[1], "_", estrato)]] <- resultado$estadisticas_100
    resultados_estadisticos_multi2cong[[paste0(variable_interes, "_", mes_subset$Mes[1], "_", estrato, "_200")]] <- resultado$estadisticas_200
    
    simulaciones_100 <- data.frame(Valor_Simulado = resultado$resultados_100) %>%
      mutate(Tamano_Muestra = 100, Variable = variable_interes, Mes = mes_subset$Mes[1], Estrato = estrato, Conglomerado = conglomerado)
    simulaciones_200 <- data.frame(Valor_Simulado = resultado$resultados_200) %>%
      mutate(Tamano_Muestra = 200, Variable = variable_interes, Mes = mes_subset$Mes[1], Estrato = estrato, Conglomerado = conglomerado)
    
    resultados_simulaciones_multi2cong[[paste0(variable_interes, "_", mes_subset$Mes[1], "_", estrato)]] <- simulaciones_100
    resultados_simulaciones_multi2cong[[paste0(variable_interes, "_", mes_subset$Mes[1], "_", estrato, "_200")]] <- simulaciones_200
  }
}

# 4. Generación de Archivos de Salida

## 4.1 Consolidar todas las estadísticas en un solo data.frame
estadisticas_completas <- do.call(rbind, resultados_estadisticos_multi2cong)

## 4.2 Consolidar todas las simulaciones en un solo data.frame
simulaciones_completas <- do.call(rbind, resultados_simulaciones_multi2cong)

## 4.3 Crear el workbook
wb <- createWorkbook()

## 4.4 Guardar estadísticas en una sola hoja
addWorksheet(wb, "Estadisticas_Multi2Cong")
writeData(wb, "Estadisticas_Multi2Cong", estadisticas_completas)

## 4.5 Guardar simulaciones en una sola hoja
addWorksheet(wb, "Simulaciones_Multi2Cong")
writeData(wb, "Simulaciones_Multi2Cong", simulaciones_completas)

## 4.6 Guardar el workbook en el archivo especificado
saveWorkbook(wb, file.path(directorio_salida, "Resultados_multi2cong.xlsx"), overwrite = TRUE)

print("Resultados guardados exitosamente en el archivo especificado.")

# 5. Generación de Gráficos

## 5.1 Función para crear gráficos de las distribuciones de las estimaciones por estrato y tamaño de muestra
crear_graficos_distribuciones_multi2cong <- function(simulaciones, valor_real) {
  ggplot(simulaciones, aes(x = Valor_Simulado)) +
    geom_histogram(binwidth = 5, fill = "blue", color = "black", alpha = 0.7) +
    geom_vline(xintercept = valor_real, color = "red", linetype = "dashed", size = 1) +
    facet_grid(Estrato + Conglomerado ~ Tamano_Muestra + Mes, scales = "free") +
    labs(
      title = "Distribución de Estimaciones por Estrato, Conglomerado y Tamaño de Muestra",
      x = "Valor Simulado",
      y = "Frecuencia"
    ) +
    theme_minimal()
}

## 5.2 Generar gráficos para cada combinación de estrato y mes
graficos <- list()

for (subset_name in names(meses_subsets)) {
  simulaciones_subset <- simulaciones_completas %>% filter(Mes %in% meses_subsets[[subset_name]]$Mes)
  valor_real <- mean(simulaciones_subset$Valor_Simulado)  # Aquí asumimos que el valor real es la media de las simulaciones
  
  grafico <- crear_graficos_distribuciones_multi2cong(simulaciones_subset, valor_real)
  graficos[[subset_name]] <- grafico
  
  # Guardar el gráfico
  ggsave(filename = file.path(directorio_salida, paste0("Grafico_Distribucion_Multi2Cong_", subset_name, ".png")), plot = grafico, width = 16, height = 12)
}

## 5.3 Mostrar gráficos en el entorno de trabajo (si se desea visualizar inmediatamente)
for (grafico in graficos) {
  print(grafico)
}

print("Gráficos de distribuciones para muestreo multietápico por conglomerados generados y guardados exitosamente.")

################################################################################
#################################multietapico2est###########################
################################################################################

# 1. Funciones

## 1.1 Función para realizar el muestreo multietápico de dos etapas utilizando dos variables de estratos
realizar_muestreo_multietapico2est <- function(datos, n_estrato, variable_interes, estrato1, estrato2) {
  combinaciones_estratos <- unique(datos %>% select(!!sym(estrato1), !!sym(estrato2)))
  muestra_total <- NULL
  
  for (i in 1:nrow(combinaciones_estratos)) {
    datos_combinacion <- datos %>%
      filter((!!sym(estrato1) == combinaciones_estratos[[estrato1]][i]) & 
               (!!sym(estrato2) == combinaciones_estratos[[estrato2]][i]))
    
    tamanos_muestra_combinacion <- round((n_estrato * nrow(datos_combinacion)) / nrow(datos))
    
    muestra_combinacion <- datos_combinacion %>%
      sample_n(size = tamanos_muestra_combinacion, replace = FALSE)
    
    muestra_total <- rbind(muestra_total, muestra_combinacion)
  }
  
  mean(muestra_total[[variable_interes]])
}

## 1.2 Función para calcular estadísticas
calcular_estadisticas2est <- function(resultados, n, valor_verdadero, estrato1, estrato2, mes, varianza_simple, variable) {
  media_estimaciones <- mean(resultados)
  sesgo <- media_estimaciones - valor_verdadero
  varianza <- var(resultados)
  error_cuadratico_medio <- mean((resultados - valor_verdadero)^2)
  
  # Calcular efecto de diseño
  efecto_diseno <- varianza / varianza_simple
  
  data.frame(
    Estrato1 = estrato1,
    Estrato2 = estrato2,
    Mes = mes,
    Variable = variable,
    Tamano_Muestra = n,
    Media_Estimaciones = media_estimaciones,
    Sesgo = sesgo,
    Varianza = varianza,
    ECM = error_cuadratico_medio,
    Efecto_Diseno = efecto_diseno
  )
}

## 1.3 Función para procesar el muestreo multietápico
procesar_muestreo_multietapico2est <- function(mes_subset, estrato1, estrato2, variable, varianza_simple) {
  valor_verdadero <- mean(mes_subset[[variable]])
  mes <- unique(mes_subset$Mes)
  
  resultados_100 <- replicate(1000, realizar_muestreo_multietapico2est(mes_subset, 100, variable, estrato1, estrato2))
  resultados_200 <- replicate(1000, realizar_muestreo_multietapico2est(mes_subset, 200, variable, estrato1, estrato2))
  
  estadisticas_100 <- calcular_estadisticas2est(resultados_100, 100, valor_verdadero, estrato1, estrato2, mes, varianza_simple, variable)
  estadisticas_200 <- calcular_estadisticas2est(resultados_200, 200, valor_verdadero, estrato1, estrato2, mes, varianza_simple, variable)
  
  list(
    estadisticas_100 = estadisticas_100,
    estadisticas_200 = estadisticas_200,
    resultados_100 = resultados_100,
    resultados_200 = resultados_200
  )
}

# 2. Configuración de estratos y variables de interés

## 2.1 Definición de estratos y variable de interés
estratos <- list(
  c("Cod_Region", "Nivel_Atencion"),      
  c("Cod_Region", "Tipo_Establec"),       
  c("Cod_Region", "Servicio_Salud"),      
  c("Servicio_Salud", "Nivel_Atencion"),  
  c("Servicio_Salud", "Tipo_Establec"),   
  c("Nivel_Atencion", "Tipo_Establec")    
)

variable_interes <- "TotalOrganizaciones"
meses_subsets <- list(
  maxvar_organizaciones = maxvar_organizaciones,
  promvar_organizaciones = promvar_organizaciones
)

resultados_estadisticas_multietapico2est <- list()
resultados_simulaciones_multietapico2est <- list()

## 2.2 Obtener varianza del muestreo aleatorio simple para cada subconjunto de mes
varianza_simple_list <- sapply(meses_subsets, function(mes_subset) {
  var(replicate(1000, realizar_muestreo_aleatorio_simple(mes_subset, 100, variable_interes)))
})

# 3. Procesar todas las combinaciones de estratos

for (subset_name in names(meses_subsets)) {
  mes_subset <- meses_subsets[[subset_name]]
  for (estrato_comb in estratos) {
    estrato1 <- estrato_comb[1]
    estrato2 <- estrato_comb[2]
    varianza_simple <- varianza_simple_list[subset_name]
    resultado <- procesar_muestreo_multietapico2est(mes_subset, estrato1, estrato2, variable_interes, varianza_simple)
    
    # Almacenar resultados
    resultados_estadisticas_multietapico2est[[paste0(variable_interes, "_", mes_subset$Mes[1], "_", estrato1, "_", estrato2, "_100")]] <- resultado$estadisticas_100
    resultados_estadisticas_multietapico2est[[paste0(variable_interes, "_", mes_subset$Mes[1], "_", estrato1, "_", estrato2, "_200")]] <- resultado$estadisticas_200
    
    # Almacenar simulaciones con identificación
    simulaciones_100 <- data.frame(Valor_Simulado = resultado$resultados_100) %>%
      mutate(Tamano_Muestra = 100, Variable = variable_interes, Mes = mes_subset$Mes[1], Estrato1 = estrato1, Estrato2 = estrato2)
    
    simulaciones_200 <- data.frame(Valor_Simulado = resultado$resultados_200) %>%
      mutate(Tamano_Muestra = 200, Variable = variable_interes, Mes = mes_subset$Mes[1], Estrato1 = estrato1, Estrato2 = estrato2)
    
    resultados_simulaciones_multietapico2est[[paste0(variable_interes, "_", mes_subset$Mes[1], "_", estrato1, "_", estrato2, "_100")]] <- simulaciones_100
    resultados_simulaciones_multietapico2est[[paste0(variable_interes, "_", mes_subset$Mes[1], "_", estrato1, "_", estrato2, "_200")]] <- simulaciones_200
  }
}

# 4. Combinar todos los resultados en un solo data.frame

## 4.1 Combinar resultados de estadísticas y simulaciones
estadisticas_completas2est <- do.call(rbind, resultados_estadisticas_multietapico2est)
simulaciones_completas2est <- do.call(rbind, resultados_simulaciones_multietapico2est)

# 5. Crear un archivo Excel con una sola hoja para los resultados numéricos

## 5.1 Crear el workbook
wb <- createWorkbook()

## 5.2 Guardar estadísticas en una hoja
addWorksheet(wb, "Estadisticas_Multietapico2est")
writeData(wb, "Estadisticas_Multietapico2est", estadisticas_completas2est)

## 5.3 Guardar simulaciones en otra hoja
addWorksheet(wb, "Simulaciones_Multietapico2est")
writeData(wb, "Simulaciones_Multietapico2est", simulaciones_completas2est)

## 5.4 Guardar el workbook en el archivo especificado
saveWorkbook(wb, file.path(directorio_salida, "Resultados_multietapico2est.xlsx"), overwrite = TRUE)

print("Resultados guardados exitosamente en el archivo especificado.")

# 6. Generación de Gráficos

## 6.1 Función para crear gráficos de las distribuciones de las estimaciones por combinación de estratos, tamaño de muestra y subset
crear_graficos_distribuciones_multietapico2est <- function(simulaciones, valor_real) {
  ggplot(simulaciones, aes(x = Valor_Simulado)) +
    geom_histogram(binwidth = 5, fill = "blue", color = "black", alpha = 0.7) +
    geom_vline(xintercept = valor_real, color = "red", linetype = "dashed", size = 1) +
    facet_grid(Estrato1 + Estrato2 ~ Tamano_Muestra + Mes, scales = "free") +
    labs(
      title = "Distribución de Estimaciones por Combinación de Estratos, Tamaño de Muestra y Mes",
      x = "Valor Simulado",
      y = "Frecuencia"
    ) +
    theme_minimal()
}

## 6.2 Generar gráficos para cada combinación de estratos y mes
graficos <- list()

for (subset_name in names(meses_subsets)) {
  simulaciones_subset <- simulaciones_completas2est %>% filter(Mes %in% meses_subsets[[subset_name]]$Mes)
  valor_real <- mean(simulaciones_subset$Valor_Simulado)  # Aquí asumimos que el valor real es la media de las simulaciones
  
  grafico <- crear_graficos_distribuciones_multietapico2est(simulaciones_subset, valor_real)
  graficos[[subset_name]] <- grafico
  
  # Guardar el gráfico
  ggsave(filename = file.path(directorio_salida, paste0("Grafico_Distribucion_Multietapico2est_", subset_name, ".png")), plot = grafico, width = 16, height = 12)
}

## 6.3 Mostrar gráficos en el entorno de trabajo (si se desea visualizar inmediatamente)
for (grafico in graficos) {
  print(grafico)
}

print("Gráficos de distribuciones para muestreo multietápico por combinación de estratos generados y guardados exitosamente.")

################################################################################
#########################ANALISIS_JERARQUICO####################################
################################################################################

# Crear un nuevo workbook para consolidar todos los resultados
wb <- createWorkbook()

# Función para agregar y consolidar resultados, con conversión de Mes a carácter
agregar_hojas_y_consolidar <- function(resultados, nombre_hoja, wb) {
  if (length(resultados) > 0) {
    estadisticas_completas <- do.call(rbind, resultados)
    
    # Convertir la columna Mes a carácter para evitar conflictos al combinar
    if ("Mes" %in% colnames(estadisticas_completas)) {
      estadisticas_completas$Mes <- as.character(estadisticas_completas$Mes)
    }
    
    addWorksheet(wb, nombre_hoja)
    writeData(wb, nombre_hoja, estadisticas_completas)
    return(estadisticas_completas)
  } else {
    warning(paste("La sección", nombre_hoja, "está vacía y no se ha agregado al workbook."))
    return(NULL)
  }
}

# Consolidar y agregar resultados de las diferentes secciones
secciones <- list(
  Aleatorio = resultados_estadisticas, # Resultados del muestreo aleatorio simple
  Estratificado = resultados_estadisticas_estratificado, # Resultados del muestreo estratificado
  Conglomerados = resultados_estadisticas_conglomerados, # Resultados del muestreo por conglomerados
  Multiestcong = resultados_estadisticos_multi2cong, # Resultados del muestreo multietápico con conglomerados
  Multietapico_2Est = resultados_estadisticas_multietapico2est # Resultados del muestreo multietápico con dos estratos
)

# Aplicar la consolidación a todas las secciones, eliminando NULL si hay secciones vacías
todas_estadisticas <- bind_rows(lapply(names(secciones), function(nombre) {
  agregar_hojas_y_consolidar(secciones[[nombre]], nombre, wb)
}) %>% discard(is.null))

# Guardar las estadísticas consolidadas en una hoja
addWorksheet(wb, "Todas_Estadisticas")
writeData(wb, "Todas_Estadisticas", todas_estadisticas)

# Análisis Jerárquico (Asignación de Pesos)
# Define los pesos para cada criterio (sin incluir Media_Estimaciones)
pesos <- c(
  Sesgo = 0.3,
  Varianza = 0.25,
  ECM = 0.3,
  Efecto_Diseno = 0.15
)

# Calcular puntuación ponderada para cada fila
todas_estadisticas <- todas_estadisticas %>%
  rowwise() %>%
  mutate(
    Puntuacion_Total = sum(
      abs(Sesgo) * pesos["Sesgo"],  # Abs es usado para que los sesgos negativos también se penalicen
      Varianza * pesos["Varianza"],
      ECM * pesos["ECM"],
      Efecto_Diseno * pesos["Efecto_Diseno"]
    )
  ) %>%
  ungroup()

# Ordenar los resultados según la puntuación total
todas_estadisticas_ordenadas <- todas_estadisticas %>%
  arrange(Puntuacion_Total)

# Guardar los resultados ordenados en una nueva hoja de Excel
addWorksheet(wb, "Resultados_Ordenados")
writeData(wb, "Resultados_Ordenados", todas_estadisticas_ordenadas)

# Guardar el archivo Excel final
saveWorkbook(wb, file.path(directorio_salida, "Resultados_Consolidados_Ordenados.xlsx"), overwrite = TRUE)

print("Resultados consolidados y ordenados guardados exitosamente en el directorio especificado.")



