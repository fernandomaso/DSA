##################################################################################
#                 INSTALAÇÃO E CARREGAMENTO DE PACOTES NECESSÁRIOS               #
##################################################################################
#Pacotes utilizados
pacotes <- c("plotly","tidyverse","ggrepel","fastDummies","knitr","kableExtra",
             "splines","reshape2","PerformanceAnalytics","correlation","see",
             "ggraph","psych","nortest","rgl","car","ggside","tidyquant","olsrr",
             "jtools","ggstance","magick","cowplot","emojifont","beepr", "Rcpp",
             "writexl", "readxl", "misc3d", "writexl", "plot3D", "cluster", "factoextra", "ade4")

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

#Abrir arquivo de Reclamacoes 2021/2022
Tabela_Reclamacoes <- read_excel("Banco_Dados_Reclamacoes_PPM.xlsx")

#Gerar Recortes: todas as combinações válidades de Janelas, Países e Agrupamentos
Parametros_Loop <- unique(select(Tabela_Reclamacoes, Janela_ID, Pais_ID, Agrupamento_ID))
#Ordernar para facilitar a visualização posteriormente
Parametros_Loop[
  with(Parametros_Loop, order(Janela_ID, Pais_ID, Agrupamento_ID),),
]

#Loop para percorrer todos os Recortes: Todas as janelas, países e agrupamentos
for (i in c(1:nrow(Parametros_Loop))) {
  #_______________________________RECORTAR O I-ÉSSIMO RECORTE_________________________________________#
  #O For Vai percorrer cada linha da Tabela "Parametros_Loop" com todas as combinações de Janela, País e Agrupamento possíveis
  #Identificar a Janela da i-éssima linha:
  Janela <- as.numeric(Parametros_Loop[as.numeric(i),1])
  #Identificar o país da i-ésima linha:
  Pais <- as.numeric(Parametros_Loop[as.numeric(i),2])
  #Identificar o agrupamento da i-éssima linha:
  Agrupamento <- as.numeric(Parametros_Loop[as.numeric(i),3])
  #Recortar a tabela do i-ésimo objeto:
  Tabela_Filtrada_i <- Tabela_Reclamacoes[Tabela_Reclamacoes$Janela_ID==Janela & Tabela_Reclamacoes$Pais_ID==Pais & Tabela_Reclamacoes$Agrupamento_ID==Agrupamento,]
  #_________________________________________________________________________________________________#

  #______________ESTIMAR MODELOS POLINOMIAIS (GRAU 1, 2 e 3) PARA O I-ÉSIMO RECORTE_________________#
  #Estimação de modelo polinomial grau 1
  Modelo_Poly_Linear_i <- lm(formula = PPM_Rec_Semana ~ poly(Tabela_Filtrada_i$Semana_ID,degree=1, raw=TRUE), data = Tabela_Filtrada_i)
  #Estimação de modelo polinomial grau 2
  Modelo_Poly_Quad_i <- lm(formula = PPM_Rec_Semana ~ poly(Tabela_Filtrada_i$Semana_ID, degree=2, raw=TRUE), data = Tabela_Filtrada_i)
  #Estimação de modelo polinomial grau 3
  modelo_Poly_Cubic_i <- lm(formula = PPM_Rec_Semana ~ poly(Tabela_Filtrada_i$Semana_ID, degree=3, raw=TRUE), data = Tabela_Filtrada_i)
  #_________________________________________________________________________________________________#

  #__________ESCOLHER O MELHOR MODELO POLINOMIAL (SEM MODELO, GRAU 1, GRAU 2 OU GRAU 3)_____________#
  #Verificar se o Beta do modelo linear passa no teste estatístico à 95% (t-student) e se o modelo passa nos testes de verificação da Aderência dos resíduos à normalidade (Shapiro Francia e Shapiro Wilk)
  if (summary(Modelo_Poly_Linear_i)$coefficients[2,4] < 0.05 & shapiro.test(Modelo_Poly_Linear_i$residuals)[2]>0.05 & sf.test(Modelo_Poly_Linear_i$residuals)[2]>0.05){
    Teste_Estatisco_Linear_i <- TRUE
    Grau_Modelo_Poly_i <- 1
  } else {
    Teste_Estatisco_Linear_i <- FALSE
    Grau_Modelo_Poly_i <- 0
  }
  #Verificar se os Betas do modelo polinomial quadrático passam no teste estatístico à 95% (t-student) e se o modelo passa nos testes de verificação da Aderência dos resíduos à normalidade (Shapiro Francia e Shapiro Wilk)
  if (summary(Modelo_Poly_Quad_i)$coefficients[2,4] < 0.05 & summary(Modelo_Poly_Quad_i)$coefficients[3,4] < 0.05 & shapiro.test(Modelo_Poly_Quad_i$residuals)[2]>0.05 & sf.test(Modelo_Poly_Quad_i$residuals)[2]>0.05){
    Teste_Estatisco_Quad_i <- TRUE
    #Comparar o modelo quadrático com o modelo linear através do comando anova para verificar se é superior estatisticamente (Caso o p-value <0,05)
    if (anova(Modelo_Poly_Linear_i, Modelo_Poly_Quad_i)$`Pr(>F)`[2]<0.05) {
      #Elevar o grau do polinomio escolhido
      Grau_Modelo_Poly_i <- 2
    }
  } else {
    #Não modificar o grau do polinomio escolhido (Matendo o do passo anterior)
    Teste_Estatisco_Quad_i <- FALSE
  }
  #Verificar se os Betas do modelo polinomial cúbico passam no teste estatístico à 95% (t-student) e se o modelo passa nos testes de verificação da Aderência dos resíduos à normalidade (Shapiro Francia e Shapiro Wilk)
  if (summary(modelo_Poly_Cubic_i)$coefficients[2,4] < 0.05 & summary(modelo_Poly_Cubic_i)$coefficients[3,4] < 0.05 & summary(modelo_Poly_Cubic_i)$coefficients[4,4] < 0.05 & shapiro.test(modelo_Poly_Cubic_i$residuals)[2]>0.05 & sf.test(modelo_Poly_Cubic_i$residuals)[2]>0.05){
    Teste_Estatisco_Cubic_i <- TRUE
    #Comparar o modelo cúbico com o modelo quadrático através do comando anova para verificar se é superior estatisticamente (Caso o p-value <0,05)
    if (anova(Modelo_Poly_Quad_i, modelo_Poly_Cubic_i)$`Pr(>F)`[2]<0.05) {
      #Elevar o grau do polinomio escolhido
      Grau_Modelo_Poly_i <- 3
    }
  } else {
    #Não modificar o grau do polinomio escolhido (Matendo o do passo anterior)
    Teste_Estatisco_Cubic_i <- FALSE
  }
  #_________________________________________________________________________________________________#
  
  #_________________________________________________________________________________________________#
  #Verificar se os quatro primeiros pontos são zero. Caso positivo forçar grau=0 porque será melhor
  #avaliar este agrupamento na próxima Janela (4 semanas a frente). Isso evitará erros de classificação (Falso positivo).
  if(Tabela_Filtrada_i$PPM_Rec_Semana[1]==0 & Tabela_Filtrada_i$PPM_Rec_Semana[2]==0 & Tabela_Filtrada_i$PPM_Rec_Semana[3]==0 & Tabela_Filtrada_i$PPM_Rec_Semana[4]==0){
    Grau_Modelo_Poly_i <- 0
  }
  #_________________________________________________________________________________________________#

  #_________________________GUARDAR VALORES INICIAIS EM TABELA RESUMO_______________________________#
  #Criar Tabela Resumo para a primeira linha
  if(i==1){
    #Criar Tabela Resumo para a primeira linha
    Tabela_Resumo <- select(Tabela_Filtrada_i, Pais_Nome, Pais_ID, Agrupamento_Nome, Agrupamento_ID, Janela_ID)[1,]
    Tabela_Resumo$PPM_Rec_Semana <- NA
    Tabela_Resumo$Semana_Inicial <- 0
    Tabela_Resumo$Semana_Final <- 0
    Tabela_Resumo$Grau_Modelo_Poly <- 0
    Tabela_Resumo$Alpha_Intercept <- 0
    Tabela_Resumo$Alpha_Int_P_Value <- 0
    Tabela_Resumo$Beta_Poly_1 <- 0
    Tabela_Resumo$Beta_Poly_1_P_Value <- 0
    Tabela_Resumo$Beta_Poly_2 <- 0
    Tabela_Resumo$Beta_Poly_2_P_Value <- 0
    Tabela_Resumo$Beta_Poly_3 <- 0
    Tabela_Resumo$Beta_Poly_3_P_Value <- 0
    Tabela_Resumo$Fitted_Values_Poly <- NA
    Tabela_Resumo$R_Quadrado_Poly <- 0
    Tabela_Resumo$Teste_Shapiro_Francia <- 0
    Tabela_Resumo$Teste_Shapiro_Walk <- 0
    Tabela_Resumo$Tendencia_Geral_Pontos <- 0
    Tabela_Resumo$Fitted_PPM_Rec_Ultima_Semana <- 0
    Tabela_Resumo$Derivada_Primeira_Ponto_Max <- 0
    Tabela_Resumo$Derivada_Primeira <- 0
    Tabela_Resumo$Derivada_Segunda <- 0
  }

  #Agregar na Tabela Resumo a i-ésima linha
  Tabela_Resumo_i <- select(Tabela_Filtrada_i, Pais_Nome, Pais_ID, Agrupamento_Nome, Agrupamento_ID, Janela_ID)[1,]
  Tabela_Resumo_i$PPM_Rec_Semana <- toString(Tabela_Filtrada_i$PPM_Rec_Semana)
  Tabela_Resumo_i$Semana_Inicial <- min(Tabela_Filtrada_i$Semana_ID)
  Tabela_Resumo_i$Semana_Final <- max(Tabela_Filtrada_i$Semana_ID)
  Tabela_Resumo_i$Grau_Modelo_Poly <- Grau_Modelo_Poly_i

  #Guardar valores específicos à depender do grau do modelo escolhido
  switch(   
    toString(Grau_Modelo_Poly_i),
    #Caso não exista modelo:
    "0"= {
      #Guardar o Alpha e o seu respectivo p-value (é nulo por não ter modelo)
      Tabela_Resumo_i$Alpha_Intercept <- NA
      Tabela_Resumo_i$Alpha_Int_P_Value <- NA
      #Guardar os valores de Beta 1 e o seu P-Value
      Tabela_Resumo_i$Beta_Poly_1 <- NA
      Tabela_Resumo_i$Beta_Poly_1_P_Value <- NA
      #Guardar os valores de Beta 2 e o seu P-Value
      Tabela_Resumo_i$Beta_Poly_2 <- NA
      Tabela_Resumo_i$Beta_Poly_2_P_Value <- NA
      #Guardar os valores de Beta 3 e o seu P-Value
      Tabela_Resumo_i$Beta_Poly_3 <- NA
      Tabela_Resumo_i$Beta_Poly_3_P_Value <- NA
      #Guardar os fitted values
      Tabela_Resumo_i$Fitted_Values_Poly <- NA
      #Guardar o valor de R^2 do modelo nulo
      Tabela_Resumo_i$R_Quadrado_Poly <- NA
      #Guardar o teste de Shapiro Francia
      Tabela_Resumo_i$Teste_Shapiro_Francia <- NA
      #Guardar o teste de Shapiro Walk
      Tabela_Resumo_i$Teste_Shapiro_Walk <- NA
      #Guardar a tendência geral dos pontos
      Tabela_Resumo_i$Tendencia_Geral_Pontos <- 0
      #Guardar o Fitted Value de PPM de Reclamacões do último ponto
      Tabela_Resumo_i$Fitted_PPM_Rec_Ultima_Semana <- 0
      #Derivadas são nulas (não fazem sentido) caso nenhum modelo válido seja encontrado
      Tabela_Resumo_i$Derivada_Primeira_Ponto_Max <- 0
      Tabela_Resumo_i$Derivada_Primeira <- 0
      Tabela_Resumo_i$Derivada_Segunda <- 0
    },
    #Caso o melhor modelo seja o linear
    "1"= {
      #Guardar o Alpha e o seu respectivo p-value
      Tabela_Resumo_i$Alpha_Intercept <- Modelo_Poly_Linear_i$coefficients[1]
      Tabela_Resumo_i$Alpha_Int_P_Value <- summary(Modelo_Poly_Linear_i)$coefficients[1,4]
      #Guardar os valores de Beta 1 e o seu P-Value
      Tabela_Resumo_i$Beta_Poly_1 <- Modelo_Poly_Linear_i$coefficients[2]
      Tabela_Resumo_i$Beta_Poly_1_P_Value <- summary(Modelo_Poly_Linear_i)$coefficients[2,4]
      #Guardar os valores de Beta 2 e o seu P-Value
      Tabela_Resumo_i$Beta_Poly_2 <- NA
      Tabela_Resumo_i$Beta_Poly_2_P_Value <- NA
      #Guardar os valores de Beta 3 e o seu P-Value
      Tabela_Resumo_i$Beta_Poly_3 <- NA
      Tabela_Resumo_i$Beta_Poly_3_P_Value <- NA
      #Guardar os fitted values
      Tabela_Resumo_i$Fitted_Values_Poly <- toString(round(Modelo_Poly_Linear_i$fitted.values,1))
      #Guardar o valor de R^2 do modelo linear
      Tabela_Resumo_i$R_Quadrado_Poly <- round(summary(Modelo_Poly_Linear_i)$r.squared, 4)
      #Guardar o teste de Shapiro Francia
      Tabela_Resumo_i$Teste_Shapiro_Francia <- as.numeric(sf.test(Modelo_Poly_Linear_i$residuals)[2])
      #Guardar o teste de Shapiro Walk
      Tabela_Resumo_i$Teste_Shapiro_Walk <- as.numeric(shapiro.test(Modelo_Poly_Linear_i$residuals)[2])
      #Guardar a tendência geral dos pontos
      Tabela_Resumo_i$Tendencia_Geral_Pontos <- Modelo_Poly_Linear_i$coefficients[2]
      #Guardar o Fitted Value de PPM de Reclamacões do último ponto
      Tabela_Resumo_i$Fitted_PPM_Rec_Ultima_Semana <- Modelo_Poly_Linear_i$fitted.values[12]
      #Encontrar o ponto de máximo e guardar o valor da derivada desse ponto. No caso do linear a derivada será a mesma em qualquer ponto
      Tabela_Resumo_i$Derivada_Primeira_Ponto_Max <- Modelo_Poly_Linear_i$coefficients[2]
      #Para Y=ax+b a primeira derivada equivale a "a" e a segunda derivada é zero
      Tabela_Resumo_i$Derivada_Primeira <- Modelo_Poly_Linear_i$coefficients[2]
      Tabela_Resumo_i$Derivada_Segunda <- 0
    }, 
    #Caso o melhor modelo seja o quadrático
    "2"= {
      #Guardar o Alpha e o seu respectivo p-value
      Tabela_Resumo_i$Alpha_Intercept <- Modelo_Poly_Quad_i$coefficients[1]
      Tabela_Resumo_i$Alpha_Int_P_Value <- summary(Modelo_Poly_Quad_i)$coefficients[1,4]
      #Guardar os valores de Beta 1 e o seu P-Value
      Tabela_Resumo_i$Beta_Poly_1 <- Modelo_Poly_Quad_i$coefficients[2]
      Tabela_Resumo_i$Beta_Poly_1_P_Value <- summary(Modelo_Poly_Quad_i)$coefficients[2,4]
      #Guardar os valores de Beta 2 e o seu P-Value
      Tabela_Resumo_i$Beta_Poly_2 <- Modelo_Poly_Quad_i$coefficients[3]
      Tabela_Resumo_i$Beta_Poly_2_P_Value <- summary(Modelo_Poly_Quad_i)$coefficients[3,4]
      #Guardar os valores de Beta 3 e o seu P-Value
      Tabela_Resumo_i$Beta_Poly_3 <- NA
      Tabela_Resumo_i$Beta_Poly_3_P_Value <- NA
      #Guardar os fitted values
      Tabela_Resumo_i$Fitted_Values_Poly <- toString(round(Modelo_Poly_Quad_i$fitted.values,1))
      #Guardar o valor de R^2 do modelo Quadrático
      Tabela_Resumo_i$R_Quadrado_Poly <- round(summary(Modelo_Poly_Quad_i)$r.squared, 4)
      #Guardar o teste de Shapiro Francia
      Tabela_Resumo_i$Teste_Shapiro_Francia <- as.numeric(sf.test(Modelo_Poly_Quad_i$residuals)[2])
      #Guardar o teste de Shapiro Walk
      Tabela_Resumo_i$Teste_Shapiro_Walk <- as.numeric(shapiro.test(Modelo_Poly_Quad_i$residuals)[2])
      #Guardar a tendência geral dos pontos
      Tabela_Resumo_i$Tendencia_Geral_Pontos <- Modelo_Poly_Linear_i$coefficients[2]
      #Guardar o Fitted Value de PPM de Reclamacões do último ponto
      Tabela_Resumo_i$Fitted_PPM_Rec_Ultima_Semana <- Modelo_Poly_Quad_i$fitted.values[12]
      #Encontrar o ponto de máximo e guardar o valor da derivada desse ponto. Para quadrático será: 2*a*x_max+b
      Tabela_Resumo_i$Derivada_Primeira_Ponto_Max <- 2*Modelo_Poly_Quad_i$coefficients[3]*Tabela_Filtrada_i$Semana_ID[which.max(Tabela_Filtrada_i$PPM_Rec_Semana)]+Modelo_Poly_Quad_i$coefficients[2]
      #Para Y=ax^2+bx+c a primeira derivada equivale a 2ax+b e a segunda derivada equivale a 2a
      Tabela_Resumo_i$Derivada_Primeira <- 2*Modelo_Poly_Quad_i$coefficients[3]*max(Tabela_Filtrada_i$Semana_ID)+Modelo_Poly_Quad_i$coefficients[2]
      Tabela_Resumo_i$Derivada_Segunda <- 2*Modelo_Poly_Quad_i$coefficients[3]
    },  
    #Caso o melhor modelo seja o cúbico
    "3"= {
      #Guardar o Alpha e o seu respectivo p-value
      Tabela_Resumo_i$Alpha_Intercept <- modelo_Poly_Cubic_i$coefficients[1]
      Tabela_Resumo_i$Aplha_Int_P_Value <- summary(modelo_Poly_Cubic_i)$coefficients[1,4]
      #Guardar os valores de Beta 1 e o seu P-Value
      Tabela_Resumo_i$Beta_Poly_1 <- modelo_Poly_Cubic_i$coefficients[2]
      Tabela_Resumo_i$Beta_Poly_1_P_Value <- summary(modelo_Poly_Cubic_i)$coefficients[2,4]
      #Guardar os valores de Beta 2 e o seu P-Value
      Tabela_Resumo_i$Beta_Poly_2 <- modelo_Poly_Cubic_i$coefficients[3]
      Tabela_Resumo_i$Beta_Poly_2_P_Value <- summary(modelo_Poly_Cubic_i)$coefficients[3,4]
      #Guardar os valores de Beta 3 e o seu P-Value
      Tabela_Resumo_i$Beta_Poly_3 <- modelo_Poly_Cubic_i$coefficients[4]
      Tabela_Resumo_i$Beta_Poly_3_P_Value <- summary(modelo_Poly_Cubic_i)$coefficients[4,4]
      #Guardar os fitted values
      Tabela_Resumo_i$Fitted_Values_Poly <- toString(round(modelo_Poly_Cubic_i$fitted.values,1))
      #Guardar o valor de R^2 do modelo Quadrático
      Tabela_Resumo_i$R_Quadrado_Poly <- round(summary(modelo_Poly_Cubic_i)$r.squared, 4)
      #Guardar o teste de Shapiro Francia
      Tabela_Resumo_i$Teste_Shapiro_Francia <- as.numeric(sf.test(modelo_Poly_Cubic_i$residuals)[2])
      #Guardar o teste de Shapiro Walk
      Tabela_Resumo_i$Teste_Shapiro_Walk <- as.numeric(shapiro.test(modelo_Poly_Cubic_i$residuals)[2])
      #Guardar a tendência geral dos pontos
      Tabela_Resumo_i$Tendencia_Geral_Pontos <- Modelo_Poly_Linear_i$coefficients[2]
      #Guardar o Fitted Value de PPM de Reclamacões do último ponto
      Tabela_Resumo_i$Fitted_PPM_Rec_Ultima_Semana <- modelo_Poly_Cubic_i$fitted.values[12]
      #Encontrar o ponto de máximo e guardar o valor da derivada desse ponto. Para o modelo cúbico sera: 3*a*x_max^2+2*b*x_max+c
      Tabela_Resumo_i$Derivada_Primeira_Ponto_Max <- 3*modelo_Poly_Cubic_i$coefficients[4]*((Tabela_Filtrada_i$Semana_ID[which.max(Tabela_Filtrada_i$PPM_Rec_Semana)])^2)+2*modelo_Poly_Cubic_i$coefficients[3]*Tabela_Filtrada_i$Semana_ID[which.max(Tabela_Filtrada_i$PPM_Rec_Semana)]+modelo_Poly_Cubic_i$coefficients[2]
      #Para Y=ax^3+bx^2+cx+d a primeira derivada equivale a 3ax^2+2bx+c e a segunda derivada equivale a 6ax+2b
      Tabela_Resumo_i$Derivada_Primeira <- 3*modelo_Poly_Cubic_i$coefficients[4]*((max(Tabela_Filtrada_i$Semana_ID))^2)+2*modelo_Poly_Cubic_i$coefficients[3]*max(Tabela_Filtrada_i$Semana_ID)+modelo_Poly_Cubic_i$coefficients[2]
      Tabela_Resumo_i$Derivada_Segunda <- 6*modelo_Poly_Cubic_i$coefficients[4]*(max(Tabela_Filtrada_i$Semana_ID))+2*modelo_Poly_Cubic_i$coefficients[3]
    }
  )  

  #Empilhar na tabela resumo:
  Tabela_Resumo[i,] <- Tabela_Resumo_i[1,]
  
  #FIM DO LOOP
}
#_________________________________________________________________________________________________#


#Extrair tabela resumo:
write_xlsx(Tabela_Resumo, "C:\\Users\\Ferds\\Dropbox\\Pós Graduação\\TCC\\Modelo Poly\\Modelo Polinomial\\Tabela_Resumo.xlsx")
