##################################################################################
#                 INSTALAÇÃO E CARREGAMENTO DE PACOTES NECESSÁRIOS               #
##################################################################################
#Pacotes utilizados
pacotes <- c("plotly","tidyverse","ggrepel","fastDummies","knitr","kableExtra",
             "splines","reshape2","PerformanceAnalytics","correlation","see",
             "ggraph","psych","nortest","rgl","car","ggside","tidyquant","olsrr",
             "jtools","ggstance","magick","cowplot","emojifont","beepr", "Rcpp",
             "writexl", "readxl", "misc3d", "writexl", "plot3D", "cluster", "factoextra", "ade4", "dplyr")

options(rgl.debug = TRUE)

if(sum(as.numeric(!pacotes %in% installed.packages())) != 0){
  instalador <- pacotes[!pacotes %in% installed.packages()]
  for(i in 1:length(instalador)) {
    install.packages(instalador, dependencies = T)
    break()}
  sapply(pacotes, require, character = T) 
} else {
  sapply(pacotes, require, character = T) 
}
##################################################################################

#Abrir arquivo output do Modelo Anterior polinomial
Tabela_Resumo_Clustering <- read_excel("Tabela_Resumo.xlsx")


#Gerar tabela para modelo Clustering com os parâmetros de recorte:
Janela_Clusting <-2
Pais_Clustering <-1
Tabela_Clustering <- select((Tabela_Resumo_Clustering[Tabela_Resumo_Clustering$Janela_ID==Janela_Clusting & Tabela_Resumo_Clustering$Pais_ID==Pais_Clustering,]), Pais_Nome, Pais_ID, Agrupamento_Nome, Agrupamento_ID, Janela_ID, Grau_Modelo_Poly, Tendencia_Geral_Pontos, Derivada_Primeira, Derivada_Segunda)

# Gráfico 3D com scatter (Como temos 4 variáveis é necessário escolher 3 para visualizar no gráfico)
rownames(Tabela_Clustering) <- Tabela_Clustering$Agrupamento_ID

scatter3D(x=Tabela_Clustering$Tendencia_Geral_Pontos,
          y=Tabela_Clustering$Derivada_Primeira,
          z=Tabela_Clustering$Derivada_Segunda,
          phi = 0, bty = "g", pch = 20, cex = 2,
          xlab = "Tendência Geral dos Pontos",
          ylab = "Variação de Reclamações no tempo",
          zlab = "Variação de Velocidade de Reclamações no tempo",
          main = "Reclamações",
          clab = "PPM Reclamações")>
  text3D(x=Tabela_Clustering$Tendencia_Geral_Pontos,
         y=Tabela_Clustering$Derivada_Primeira,
         z=Tabela_Clustering$Derivada_Segunda,
         labels = rownames(Tabela_Clustering),
         add = TRUE, cex = 1)

# Estatísticas descritivas
summary(Tabela_Clustering$Derivada_Primeira)
summary(Tabela_Clustering$Derivada_Segunda)
summary(Tabela_Clustering$Tendencia_Geral_Pontos)

#______________________________________________________________________________#
#Esquema de aglomeração hierárquico
# Matriz de dissimilaridades
matriz_D <- Tabela_Clustering %>% 
  select(Tendencia_Geral_Pontos, Derivada_Primeira, Derivada_Segunda) %>% 
  dist(method = "manhattan")
# Method: parametrização da distância a ser utilizada
## "euclidean": distância euclidiana
## "euclidiana quadrática": elevar ao quadrado matriz_D (matriz_D^2)
## "maximum": distância de Chebychev;
## "manhattan": distância de Manhattan (ou distância absoluta ou bloco);
## "canberra": distância de Canberra;
## "minkowski": distância de Minkowski
#______________________________________________________________________________#

#______________________________________________________________________________#
# Elaboração da clusterização hierárquica
cluster_hier <- agnes(x = matriz_D, method = "average")
# Method é o tipo de encadeamento:
## "complete": encadeamento completo (furthest neighbor ou complete linkage)
## "single": encadeamento único (nearest neighbor ou single linkage)
## "average": encadeamento médio (between groups ou average linkage)
#______________________________________________________________________________#

#______________________________________________________________________________#
# Definição do esquema hierárquico de aglomeração
# As distâncias para as combinações em cada estágio
coeficientes <- sort(cluster_hier$height, decreasing = FALSE)
# Construção do dendrograma
dev.off()
fviz_dend(x = cluster_hier)
#______________________________________________________________________________#

#______________________________________________________________________________#
# Método de Elbow para identificação do número ótimo de clusters
fviz_nbclust(Tabela_Clustering[,7:9], kmeans, method = "wss", k.max = 10)
#______________________________________________________________________________#

#______________________________________________________________________________#
#A partir do gráfico de Elbow definir na variável abaixo o número de Clusters
Clusters <- 7
#______________________________________________________________________________#

#______________________________________________________________________________#
# Dendrograma com visualização dos clusters
fviz_dend(x = cluster_hier,
          k = Clusters,
          k_colors = c("deeppink4", "darkviolet", "deeppink"),
          color_labels_by_k = F,
          rect = T,
          rect_fill = T,
          lwd = 1,
          ggtheme = theme_bw())
#______________________________________________________________________________#


#______________________________________________________________________________#
# Criando variável categórica para indicação do cluster no banco de dados
## O argumento 'k' indica a quantidade de clusters
Tabela_Clustering$cluster_H <- factor(cutree(tree = cluster_hier, k = Clusters))
#______________________________________________________________________________#


Tabela_Clustering_Output <- group_by(Tabela_Clustering, cluster_H) %>%
  summarise(
    #Guardar o país estudado
    Pais = first(Pais_Clustering),
    #Guardar a janela estudada
    Janela = first(Janela_Clusting),
    #Contar quantos agrupamentos existem dentro do Cluster
    Count_Agrups = n_distinct(Agrupamento_ID),
    #Concatenar os agrupamentos presentes no Cluster
    Agrupamentos = paste(Agrupamento_ID, collapse = " | "),
    #Medias das variaveis
    Media_Derivada_Primeira = mean(Derivada_Primeira, na.rm = TRUE),
    Media_Derivada_Segunda = mean(Derivada_Segunda, na.rm = TRUE),
    Media_Tendencia_Geral_Pontos = mean(Tendencia_Geral_Pontos, na.rm = TRUE),
    # Estatísticas descritivas da variável 'Derivada Primeira'
    SD_Derivada_Primeira = sd(Derivada_Primeira, na.rm = TRUE),
    Min_Derivada_Primeira = min(Derivada_Primeira, na.rm = TRUE),
    Max_Derivada_Primeira = max(Derivada_Primeira, na.rm = TRUE),
    # Estatísticas descritivas da variável 'Derivada Segunda'
    SD_Derivada_Segunda = sd(Derivada_Segunda, na.rm = TRUE),
    Min_Derivada_Segunda = min(Derivada_Segunda, na.rm = TRUE),
    Max_Derivada_Segunda = max(Derivada_Segunda, na.rm = TRUE),
    # Estatísticas descritivas da variável 'Tendencial Geral dos Pontos'
    SD_Tendencia_Geral_Pontos = sd(Tendencia_Geral_Pontos, na.rm = TRUE),
    Min_Tendencia_Geral_Pontos = min(Tendencia_Geral_Pontos, na.rm = TRUE),
    Max_Tendencia_Geral_Pontos = max(Tendencia_Geral_Pontos, na.rm = TRUE))

#Extrair tabela Clustering:
write_xlsx(Tabela_Clustering_Output, "C:\\Users\\Ferds\\Dropbox\\Pós Graduação\\TCC\\Modelo Poly\\Modelo Polinomial\\Tabela_Clustering_Output.xlsx")

#Extrair tabela Clustering:
write_xlsx(Tabela_Clustering, "C:\\Users\\Ferds\\Dropbox\\Pós Graduação\\TCC\\Modelo Poly\\Modelo Polinomial\\Tabela_Clustering.xlsx")

##################CONTINUAR A PARTIR DAQUI ########################################



# A seguir, vamos verificar se todas as variáveis ajudam na formação dos grupos

summary(anova_child_mort <- aov(formula = child_mort ~ cluster_H,
                                data = pais_padronizado))

summary(anova_exports <- aov(formula = exports ~ cluster_H,
                             data = pais_padronizado))

summary(anova_health <- aov(formula = health ~ cluster_H,
                            data = pais_padronizado))

summary(anova_imports <- aov(formula = imports ~ cluster_H,
                             data = pais_padronizado))


