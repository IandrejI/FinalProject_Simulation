getwd()
setwd('C:/Users/andre/Models/SimulationsstudieSupermarkt/Data')

###############
## Library   ##
###############

library('readxl')
library('tidyverse')
library(Rmisc)
library(firatheme)
library(gridExtra)
library(grid)
###############
## Datainput ##
###############

# xlsx files
xlsx_DLZ <- read_excel('DLZ.xlsx')
xlsx_Fehlmenge <- read_excel('Fehlmenge.xlsx')
xlsx_Kundenanzahl <- read_excel('Kundenanzahl.xlsx')
xlsx_Lagermenge <- read_excel('Lagermenge.xlsx')
xlsx_Nachfragemenge <- read_excel('Nachfragemenge.xlsx')
xlsx_Trash <- read_excel('Trash.xlsx')
replications <- 300
# Cleaning
DLZ <- xlsx_DLZ[,c(seq(from = 2, to = length(xlsx_DLZ)+1, by = 2))]
DLZ <- as.data.frame(DLZ)
row.names(DLZ) <- c(xlsx_DLZ$`Rep 1 - Time`)
DLZ <- DLZ[,1:(length(DLZ)-1)]

Fehlmenge <- xlsx_Fehlmenge[,c(seq(from = 2, to = length(xlsx_Fehlmenge)+1, by = 2))]
Fehlmenge <- as.data.frame(Fehlmenge)
row.names(Fehlmenge) <- c(xlsx_Fehlmenge$`Rep 1 - Time...1`)
Fehlmenge <- Fehlmenge[,1:(length(Fehlmenge)-7)]

Kundenanzahl <- xlsx_Kundenanzahl[,c(seq(from = 2, to = length(xlsx_Kundenanzahl)+1, by = 2))]
Kundenanzahl <- as.data.frame(Kundenanzahl)
row.names(Kundenanzahl) <- c(xlsx_Kundenanzahl$`Rep 1 - Time`)
Kundenanzahl <- Kundenanzahl[,1:(length(Kundenanzahl)-1)]

Lagermenge <- xlsx_Lagermenge[,c(seq(from = 2, to = length(xlsx_Lagermenge)+1, by = 2))]
Lagermenge <- as.data.frame(Lagermenge)
row.names(Lagermenge) <- c(xlsx_Lagermenge$`Rep 1 - Time...1`)
Lagermenge <- Lagermenge[,1:(length(Lagermenge)-7)]

Nachfragemenge <- xlsx_Nachfragemenge[,c(seq(from = 2, to = length(xlsx_Nachfragemenge)+1, by = 2))]
Nachfragemenge <- as.data.frame(Nachfragemenge)
row.names(Nachfragemenge) <- c(xlsx_Nachfragemenge$`Rep 1 - Time...1`)
Nachfragemenge <- Nachfragemenge[,1:(length(Nachfragemenge)-7)]

Trash <- xlsx_Trash[,c(seq(from = 2, to = length(xlsx_Trash)+1, by = 2))]
Trash <- as.data.frame(Trash)
row.names(Trash) <- c(xlsx_Trash$`Rep 1 - Time...1`)
Trash <- Trash[,1:(length(Trash)-7)]

# Differenzen
dif_Fehlmenge <- diff(as.matrix(Fehlmenge))
dif_Fehlmenge <- as.data.frame(dif_Fehlmenge)
FehlmengeRow <- Fehlmenge[1,]
dif_Fehlmenge <- rbind(FehlmengeRow, dif_Fehlmenge)
dif_Fehlmenge$index <- as.numeric(row.names(dif_Fehlmenge))
dif_Fehlmenge <- dif_Fehlmenge[,1:(length(dif_Fehlmenge)-1)]

dif_Lagermenge <- diff(as.matrix(Lagermenge))
dif_Lagermenge <- as.data.frame(dif_Lagermenge)
LagermengeRow <- Lagermenge[1,]
dif_Lagermenge <- rbind(dif_Lagermenge, LagermengeRow)
dif_Lagermenge$index <- as.numeric(row.names(dif_Lagermenge))
dif_Lagermenge <- dif_Lagermenge[order(dif_Lagermenge$index), ]
dif_Lagermenge <- dif_Lagermenge[,1:(length(dif_Lagermenge)-1)]

dif_Nachfragemenge <- diff(as.matrix(Nachfragemenge))
dif_Nachfragemenge <- as.data.frame(dif_Nachfragemenge)
NachfragemengeRow <- Nachfragemenge[1,]
dif_Nachfragemenge <- rbind(dif_Nachfragemenge, NachfragemengeRow)
dif_Nachfragemenge$index <- as.numeric(row.names(dif_Nachfragemenge))
dif_Nachfragemenge <- dif_Nachfragemenge[order(dif_Nachfragemenge$index), ]
dif_Nachfragemenge <- dif_Nachfragemenge[,1:(length(dif_Nachfragemenge)-1)]

dif_Trash <- diff(as.matrix(Trash))
dif_Trash <- as.data.frame(dif_Trash)
TrashRow <- Trash[1,]
dif_Trash <- rbind(dif_Trash, TrashRow)
dif_Trash$index <- as.numeric(row.names(dif_Trash))
dif_Trash <- dif_Trash[order(dif_Trash$index), ]
dif_Trash <- dif_Trash[,1:(length(dif_Trash)-1)]

dif_Kundenanzahl <- diff(as.matrix(Kundenanzahl))
dif_Kundenanzahl <- as.data.frame(dif_Kundenanzahl)
KundenanzahlRow <- Kundenanzahl[1,]
dif_Kundenanzahl <- rbind(dif_Kundenanzahl, KundenanzahlRow)
dif_Kundenanzahl$index <- as.numeric(row.names(dif_Kundenanzahl))
dif_Kundenanzahl <- dif_Kundenanzahl[order(dif_Kundenanzahl$index), ]
dif_Kundenanzahl <- dif_Kundenanzahl[,1:(length(dif_Kundenanzahl)-1)]
dif_Kundenanzahl <- as.data.frame(rep(dif_Kundenanzahl, c(rep(7,replications))))

#######################
## KPI Berechnungen ##
######################

# Beta Servicegrad
## ((Nachfragemenge / Kundenanzahl) - (Fehlmenge / Kundenanzahl)) / (Nachfragemenge / Kundenanzahl)
sService <- ((dif_Nachfragemenge /  dif_Kundenanzahl) - (dif_Fehlmenge /  dif_Kundenanzahl)) / (dif_Nachfragemenge /  dif_Kundenanzahl)

## aggregiert
i <- c(seq(1,replications,1))
agg_dif_Nachfragemenge <- as.data.frame(t(aggregate(t(dif_Nachfragemenge), list(substring(names(dif_Nachfragemenge), 0, 7)), FUN=sum)))
agg_dif_Nachfragemenge[ , i] <- apply(agg_dif_Nachfragemenge[ , i], 2, function(x) as.integer(as.character(x)))
agg_dif_Fehlmenge <- as.data.frame(t(aggregate(t(dif_Fehlmenge), list(substring(names(dif_Fehlmenge), 0, 7)), FUN=sum)))
agg_dif_Fehlmenge[ , i] <- apply(agg_dif_Fehlmenge[ , i], 2, function(x) as.integer(as.character(x)))
agg_dif_Kundenanzahl <- as.data.frame(t(aggregate(t(dif_Kundenanzahl), list(substring(names(dif_Kundenanzahl), 0, 7)), FUN=sum)))
agg_dif_Kundenanzahl[ , i] <- apply(agg_dif_Kundenanzahl[ , i], 2, function(x) as.integer(as.character(x)))
sService_agg <- ((agg_dif_Nachfragemenge / agg_dif_Kundenanzahl) - (agg_dif_Fehlmenge / agg_dif_Kundenanzahl)) / (agg_dif_Nachfragemenge /  agg_dif_Kundenanzahl)
sService_mean_O1 <- sService_agg[-1,] %>% rowMeans()

## aggregiert Spalten
aggSp_Nachfrage <- sapply(agg_dif_Nachfragemenge[-1,], FUN=mean)
aggSp_Fehlmenge <- sapply(agg_dif_Fehlmenge[-1,], FUN=mean)
aggSp_Kunden <- sapply(agg_dif_Kundenanzahl[-1,], FUN=mean)



## aggregiert
agg_dif_Trash <- as.data.frame(t(aggregate(t(dif_Trash), list(substring(names(dif_Trash), 0, 7)), FUN=sum)))
agg_dif_Trash[ , i] <- apply(agg_dif_Trash[ , i], 2, function(x) as.integer(as.character(x)))
agg_dif_Lagermenge <- as.data.frame(t(aggregate(t(dif_Lagermenge), list(substring(names(dif_Lagermenge), 0, 7)), FUN=sum)))
agg_dif_Lagermenge[ , i] <- apply(agg_dif_Lagermenge[ , i], 2, function(x) as.integer(as.character(x)))
Verderbquote_agg <- agg_dif_Trash / agg_dif_Lagermenge
Verderbquote_mean <- Verderbquote_agg[-1,] %>% rowMeans(na.rm=TRUE)

## aggregiert Spalten
aggSp_Lagermenge <- sapply(agg_dif_Lagermenge[-1,], FUN=mean)
aggSp_Trash <- sapply(agg_dif_Trash[-1,], FUN=mean)



# KPIS --------------------------------------------------------------------

Verderbquote_O3 <- dif_Trash / dif_Lagermenge
Verderbquote_Spagg_O3 <- aggSp_Trash / aggSp_Lagermenge
sService_Spagg_O3 <- ((aggSp_Nachfrage / aggSp_Kunden) - (aggSp_Fehlmenge / aggSp_Kunden)) / (aggSp_Nachfrage /  aggSp_Kunden)
aggSp_DLZ_O3 <- sapply(DLZ[-1,], FUN=mean)
DLZ_mean_O3 <- DLZ %>% rowMeans()

## aggregiert Spalten


####################################################################################
#Confidence interval
#Parameter1: Vector with observations
#Parameter2: Confidence level alpha

confidence_interval<-function(sample, alpha){
  m<-mean(sample)
  v<-var(sample)
  t<-qt(p=1-(alpha/2),df=length(sample)-1)
  hir<-t*sqrt(v/length(sample))
  lowerb<-m-hir
  upperb<-m+hir
  
  plot(sample,	
       xlab='Observation',
       ylab='Value')
  
  abline(h=lowerb, col='red')
  abline(h=m, col='black')
  abline(h=upperb, col='red')
  data <- cbind(m,v,lowerb, upperb)
  return(data)
}

# Konfidenzintervalle der KPIs

##Set alpha
alpha<-0.05

##Apply
confidence_interval(as.data.frame(Verderbquote_Spagg)[,1],alpha)
confidence_interval(as.data.frame(sService_Spagg)[,1],alpha)
confidence_interval(as.data.frame(aggSp_DLZ)[,1],alpha)
confidence_interval(as.data.frame(Verderbquote_Spagg_O1)[,1],alpha)
confidence_interval(as.data.frame(sService_Spagg_O1)[,1],alpha)
confidence_interval(as.data.frame(aggSp_DLZ_O1)[,1],alpha)

######################################################################################
#Difference Test (Paired t-Test)

t_testS<-function(Verderb_f_dif, alpha){
#Calculate differences of the two samples
Dif<-Verderb_f_dif[,1]-Verderb_f_dif[,2]

#Calculate the mean of differences
m<-mean(Dif)
m

#Claculate the variance of differences
v<-var(Dif)
v

#Claculate t value with the quantile function (qt)
t<-qt(p=1-(alpha/2),df=length(Verderb_f_dif[,2])-1)
t

#Claculate the half interval range
hir<-t*sqrt(v/length(Verderb_f_dif[,2]))

#Determine the lower bound of the conf. interval by subracting the hir from the mean
lowerb<-m-hir

#Determine the upper bound of the conf. interval by adding the hir to the mean
upperb<-m+hir


#Observe if the interval covers the 0
lowerb
upperb

#Application of the t-Test in R with the "paired" specification
#For more information see: https://stat.ethz.ch/R-manual/R-devel/library/stats/html/t.test.html

a<-t.test(x=Verderb_f_dif[,1],y=Verderb_f_dif[,2], paired=TRUE,conf.level=0.95)
return(a)
}

## T Test apply
Verderb_f_dif <- cbind(Verderbquote_Spagg, Verderbquote_Spagg_O1)
t_testS(Verderb_f_dif, alpha)

Service_f_dif <- cbind(sService_Spagg, sService_Spagg_O1)
t_testS(Service_f_dif, alpha)

DLZ_f_dif <- cbind(aggSp_DLZ, aggSp_DLZ_O1)
t_testS(DLZ_f_dif, alpha)

###################################################################################
# KPIs Summary

var(Verderbquote_Spagg)
var(sService_Spagg)
var(Verderbquote_Spagg)

CI(sService_Spagg_O1)

#########################################################
dataPlot <- function(vec0, vec01, vec02, vec03){
  z <- cbind(t(CI(vec0)), prob = 0, OnlineSachbearb = 0)
 a <- cbind(t(CI(vec01)), prob = 0.1, OnlineSachbearb = 1)
 b <- cbind(t(CI(vec02)), prob = 0.2, OnlineSachbearb = 2)
 c <- cbind(t(CI(vec03)), prob = 0.3, OnlineSachbearb = 3)
 df <- as.data.frame(rbind(z,a, b, c))
 df$prob <- as.factor(df$prob)
 df$OnlineSachbearb = as.factor(df$OnlineSachbearb)
 return(df)
}
dataPlotService <- dataPlot(sService_Spagg_O, sService_Spagg_O1,sService_Spagg_O2,sService_Spagg_O3)
dataPlotVerder <- dataPlot(Verderbquote_Spagg_O,Verderbquote_Spagg_O1, Verderbquote_Spagg_O2, Verderbquote_Spagg_O3)
dataPlotDLZ <- dataPlot(aggSp_DLZ_O, aggSp_DLZ_O1,aggSp_DLZ_O2, aggSp_DLZ_O3)


(servicePlot <- ggplot(dataPlotService, aes(y = mean, x = prob, color = OnlineSachbearb)) +
  geom_point(size = 3) +
  geom_errorbar(aes(ymin = lower, ymax = upper), width = 0.25) +
    labs(y = expression(paste(beta,"-Servicegrad")),
         x = expression(p[online]),
         col = expression(N[OS])) + 
  scale_color_brewer(palette = "Set2") +
  theme_fira())

(verderbPlot <- ggplot(dataPlotVerder, aes(y = mean, x = prob, color = OnlineSachbearb)) +
  geom_point(size = 3) +
  geom_errorbar(aes(ymin = lower, ymax = upper), width = 0.25) +
  labs(y = "Verderbquote",
       x = expression(p[online]),
       col = expression(N[OS])) + 
  scale_color_brewer(palette = "Set2") +
  theme_fira())

(DLZPlot <- ggplot(dataPlotDLZ, aes(y = mean, x = prob, color = OnlineSachbearb)) +
  geom_point(size = 3) +
  geom_errorbar(aes(ymin = lower, ymax = upper), width = 0.25) +
    labs(y = "DLZ",
         x = expression(p[online]),
         col = expression(paste(N[OS]," :"))) +
  theme_fira() +
  scale_color_brewer(palette = "Set2") +
  theme(legend.position = "bottom"))

g_legend<-function(a.gplot){
  tmp <- ggplot_gtable(ggplot_build(a.gplot))
  leg <- which(sapply(tmp$grobs, function(x) x$name) == "guide-box")
  legend <- tmp$grobs[[leg]]
  return(legend)
  }
legend <- g_legend(DLZPlot)

ex <- expression(paste("KPIs in Abhängigkeit von ", p[online]," und ", N[OS]))

finalPlot <- grid.arrange(arrangeGrob(servicePlot + theme(legend.position = "none"),
                        verderbPlot + theme(legend.position = "none"),
                        DLZPlot +theme(legend.position = "none"),
                        nrow = 1), top = textGrob(ex),
                        legend, nrow=2,heights=c(10, 1))

ggsave("finalPlot.png", finalPlot, dpi = 600, width = 9, height = 5)
dev.off()
