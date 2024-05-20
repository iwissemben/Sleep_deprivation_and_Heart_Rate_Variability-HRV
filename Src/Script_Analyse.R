#-----------------------------------installation des package requis ---------------------------------------


library(chron)
library(openxlsx)
library(readxl)
library(rmarkdown)
library(tidyverse)

rm(list=ls()) #Nettoyage de l'environnement: Suppression de toutes les variables existantes

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Importation du fichier xls et preparation du dataset brut
#------------------------------------------------------------------------------------------------------------------------------------------------------------------

fichier=file.choose()#ouvrir fichier extension "xls"

#Dataframe_raw=data.frame(read_xls(fichier,col_names =TRUE))#construit l'objet dataframe a partir du dataset

Dataframe_raw = openxlsx::read.xlsx(fichier,detectDates = TRUE)
Dataframe_raw$Rmssd=as.numeric(unlist(Dataframe_raw$Rmssd))
Dataframe_raw$`PNN50(%)`=as.numeric(unlist(Dataframe_raw$`PNN50(%)`))
Dataframe_raw$SDNN=as.numeric(unlist(Dataframe_raw$SDNN))
Dataframe_raw$`LF/HF`=as.numeric(unlist(Dataframe_raw$`LF/HF`))


#summary(Dataframe_raw) #test de la bonne importation

#get all sujbect id
Unique_subject_id=unique(Dataframe_raw$Nom)
Unique_subject_id=paste(Unique_subject_id,collapse="_")

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
# ANALYSE TEMPORELLE
#------------------------------------------------------------------------------------------------------------------------------------------------------------------


#------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Graph 1 - Moyennes RMSSD des participants selon la phase et la position
#------------------------------------------------------------------------------------------------------------------------------------------------------------------

SousDF_rmssd=Dataframe_raw %>% 
  select(Date,Nom,Position,Cdt,Rmssd) %>% 
  group_by(Nom,Position,Cdt)

SousDF_rmssd_stats=SousDF_rmssd %>% 
  summarize(Moyenne.RMSSD=mean(Rmssd),
            Sd.RMSSD=sd(Rmssd),
            n.RMSSD=n(),
            SEM.RMSSD=Sd.RMSSD/sqrt(n.RMSSD))

Graph01=ggplot(data=SousDF_rmssd_stats,
               aes(x=Cdt,y=Moyenne.RMSSD,group=interaction(Nom,Position),
                   color=Nom,linetype=Position))+
  geom_line(position="identity",size=1)+
  geom_point(position="identity",size=3,color="black")+
  geom_point(data=SousDF_rmssd,aes(x=Cdt,y=Rmssd,shape=Position),size=2)+
  facet_wrap(vars(Nom),scales = "free")+
  geom_errorbar(aes(x=Cdt,
                    ymin=Moyenne.RMSSD-Sd.RMSSD,
                    ymax=Moyenne.RMSSD+Sd.RMSSD),
                size=0.8,width=0.2)+
  labs(title="Moyennes RMSSD des participants selon la phase et la position",
       subtitle="Barres d'erreur: Sd",
       x="Sujet",
       y=paste("Moyenne RMSSD (ms)"),
       caption="WBR,MS,MM")

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Graph 2 - Moyennes HR des participants selon la phase et la position
#------------------------------------------------------------------------------------------------------------------------------------------------------------------

SousDF_HR=Dataframe_raw %>% 
  select(Date,Nom,Position,Cdt,`HR.Moyen(bpm)`) %>% 
  group_by(Nom,Position,Cdt)

SousDF_HR_stats=SousDF_HR %>% 
  summarize(Moyenne.HR=mean(`HR.Moyen(bpm)`),
            Sd.HR=sd(`HR.Moyen(bpm)`),
            n.HR=n(),
            SEM.HR=Sd.HR/sqrt(n.HR))

Graph02=ggplot(data=SousDF_HR_stats,
               aes(x=Cdt,y=Moyenne.HR,group=interaction(Nom,Position),
                   color=Nom,linetype=Position))+
  geom_line(position="identity",size=1)+
  geom_point(position="identity",size=3,color="black")+
  geom_point(data=SousDF_HR,aes(x=Cdt,y=`HR.Moyen(bpm)`,shape=Position),size=2)+
  facet_wrap(vars(Nom))+
  geom_errorbar(aes(x=Cdt,
                    ymin=Moyenne.HR-Sd.HR,
                    ymax=Moyenne.HR+Sd.HR),
                size=0.8,width=0.2)+
  labs(title="Moyennes HR des participants selon la phase et la position",
       subtitle="Barres d'erreur: Sd",
       x="Sujet",
       y=paste("Moyenne HR (bpm)"),
       caption="WBR,MS,MM")

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Graph 3 - Moyennes RR des participants selon la phase et la position
#------------------------------------------------------------------------------------------------------------------------------------------------------------------

SousDF_RR=Dataframe_raw %>% 
  select(Date,Nom,Position,Cdt,`RR.moyen(ms)`) %>% 
  group_by(Nom,Position,Cdt)

SousDF_RR_stats=SousDF_RR %>% 
  summarize(Moyenne.RR=mean(`RR.moyen(ms)`),
            Sd.RR=sd(`RR.moyen(ms)`),
            n.RR=n(),
            SEM.RR=Sd.RR/sqrt(n.RR))

Graph03=ggplot(data=SousDF_RR_stats,
               aes(x=Cdt,y=Moyenne.RR,group=interaction(Nom,Position),
                   color=Nom,linetype=Position))+
  geom_line(position="identity",size=1)+
  geom_point(position="identity",size=3,color="black")+
  geom_point(data=SousDF_RR,aes(x=Cdt,y=`RR.moyen(ms)`,shape=Position),size=2)+
  facet_wrap(vars(Nom))+
  geom_errorbar(aes(x=Cdt,
                    ymin=Moyenne.RR-Sd.RR,
                    ymax=Moyenne.RR+Sd.RR),
                size=0.8,width=0.2)+
  labs(title="Moyennes RR des participants selon la phase et la position",
       subtitle="Barres d'erreur: Sd",
       x="Sujet",
       y=paste("Moyenne RR (ms)"),
       caption="WBR,MS,MM")


#------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Graph 4 - Moyennes PNN50 des participants selon la phase et la position
#------------------------------------------------------------------------------------------------------------------------------------------------------------------

SousDF_PNN50=Dataframe_raw %>% 
  select(Date,Nom,Position,Cdt,`PNN50(%)`) %>% 
  group_by(Nom,Position,Cdt)

SousDF_PNN50_stats=SousDF_PNN50 %>% 
  summarize(Moyenne.PNN50=mean(`PNN50(%)`),
            Sd.PNN50=sd(`PNN50(%)`),
            n.PNN50=n(),
            SEM.RR=Sd.PNN50/sqrt(n.PNN50))

Graph04=ggplot(data=SousDF_PNN50_stats,
               aes(x=Cdt,y=Moyenne.PNN50,group=interaction(Nom,Position),
                   color=Nom,linetype=Position))+
  geom_line(position="identity",size=1)+
  geom_point(position="identity",size=3,color="black")+
  geom_point(data=SousDF_PNN50,aes(x=Cdt,y=`PNN50(%)`,shape=Position),size=2)+
  facet_wrap(vars(Nom))+
  geom_errorbar(aes(x=Cdt,
                    ymin=Moyenne.PNN50-Sd.PNN50,
                    ymax=Moyenne.PNN50+Sd.PNN50),
                size=0.8,width=0.2)+
  labs(title="PNN50 moyens des participants selon la phase et la position",
       subtitle="Barres d'erreur: Sd",
       x="Sujet",
       y=paste("Moyenne PNN50 (%)"),
       caption="WBR,MS,MM")

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Graph 5 - Moyennes SDNN des participants selon la phase et la position
#------------------------------------------------------------------------------------------------------------------------------------------------------------------

SousDF_SDNN=Dataframe_raw %>% 
  select(Date,Nom,Position,Cdt,SDNN) %>% 
  group_by(Nom,Position,Cdt)

SousDF_SDNN_stats=SousDF_SDNN %>% 
  summarize(Moyenne.SDNN=mean(SDNN),
            Sd.SDNN=sd(SDNN),
            n.SDNN=n(),
            SEM.SDNN=Sd.SDNN/sqrt(n.SDNN))

Graph05=ggplot(data=SousDF_SDNN_stats,
               aes(x=Cdt,y=Moyenne.SDNN,group=interaction(Nom,Position),
                   color=Nom,linetype=Position))+
  geom_line(position="identity",size=1)+
  geom_point(position="identity",size=3,color="black")+
  geom_point(data=SousDF_SDNN,aes(x=Cdt,y=SDNN,shape=Position),size=2)+
  facet_wrap(vars(Nom),scales = "free")+
  geom_errorbar(aes(x=Cdt,
                    ymin=Moyenne.SDNN-Sd.SDNN,
                    ymax=Moyenne.SDNN+Sd.SDNN),
                size=0.8,width=0.2)+
  labs(title="Moyennes SDNN des participants selon la phase et la position",
       subtitle="Barres d'erreur: Sd",
       x="Sujet",
       y=paste("Moyenne SDNN (ms)"),
       caption="WBR,MS,MM")
  
#------------------------------------------------------------------------------------------------------------------------------------------------------------------
# ANALYSE Frequentielle
#------------------------------------------------------------------------------------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Graph 6 - Moyennes Puissances HF des participants selon la phase et la position
#------------------------------------------------------------------------------------------------------------------------------------------------------------------

SousDF_HF=Dataframe_raw %>% 
  select(Date,Nom,Position,Cdt,`HF(ms²)`) %>% 
  group_by(Nom,Position,Cdt)

SousDF_HF_stats=SousDF_HF %>% 
  summarize(Moyenne.HF=mean(`HF(ms²)`),
            Sd.HF=sd(`HF(ms²)`),
            n.HF=n(),
            SEM.HF=Sd.HF/sqrt(n.HF))

Graph06=ggplot(data=SousDF_HF_stats,
               aes(x=Cdt,y=Moyenne.HF,group=interaction(Nom,Position),
                   color=Nom,linetype=Position))+
  geom_line(position="identity",size=1)+
  geom_point(position="identity",size=3,color="black")+
  geom_point(data=SousDF_HF,aes(x=Cdt,y=`HF(ms²)`,shape=Position),size=2)+
  facet_wrap(vars(Nom),scales = "free")+
  geom_errorbar(aes(x=Cdt,
                    ymin=Moyenne.HF-Sd.HF,
                    ymax=Moyenne.HF+Sd.HF),
                size=0.8,width=0.2)+
  labs(title="Moyennes des puissances HF des participants selon la phase et la position",
       subtitle="Barres d'erreur: Sd",
       x="Sujet",
       y=paste("Moyenne Puissance HF (ms²)"),
       caption="WBR,MS,MM")


#------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Graph 7 - Moyennes Puissances LF des participants selon la phase et la position
#------------------------------------------------------------------------------------------------------------------------------------------------------------------

SousDF_LF=Dataframe_raw %>% 
  select(Date,Nom,Position,Cdt,`LF(ms)²`) %>% 
  group_by(Nom,Position,Cdt)

SousDF_LF_stats=SousDF_LF %>% 
  summarize(Moyenne.LF=mean(`LF(ms)²`),
            Sd.LF=sd(`LF(ms)²`),
            n.LF=n(),
            SEM.HF=Sd.LF/sqrt(n.LF))

Graph07=ggplot(data=SousDF_LF_stats,
               aes(x=Cdt,y=Moyenne.LF,group=interaction(Nom,Position),
                   color=Nom,linetype=Position))+
  geom_line(position="identity",size=1)+
  geom_point(position="identity",size=3,color="black")+
  geom_point(data=SousDF_LF,aes(x=Cdt,y=`LF(ms)²`,shape=Position),size=2)+
  facet_wrap(vars(Nom),scales = "free")+
  geom_errorbar(aes(x=Cdt,
                    ymin=Moyenne.LF-Sd.LF,
                    ymax=Moyenne.LF+Sd.LF),
                size=0.8,width=0.2)+
  labs(title="Moyennes des puissances LF des participants selon la phase et la position",
       subtitle="Barres d'erreur: Sd",
       x="Sujet",
       y=paste("Moyenne Puissance LF (ms²)"),
       caption="WBR,MS,MM")

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Graph 8 - Moyennes ratio LF/HF des participants selon la phase et la position
#------------------------------------------------------------------------------------------------------------------------------------------------------------------

SousDF_LFHF=Dataframe_raw %>% 
  select(Date,Nom,Position,Cdt,`LF/HF`) %>% 
  group_by(Nom,Position,Cdt)

SousDF_LFHF_stats=SousDF_LFHF %>% 
  summarize(Moyenne.LFHF=mean(`LF/HF`),
            Sd.LFHF=sd(`LF/HF`),
            n.LFHF=n(),
            SEM.LFHF=Sd.LFHF/sqrt(n.LFHF))

Graph08=ggplot(data=SousDF_LFHF_stats,
               aes(x=Cdt,y=Moyenne.LFHF,group=interaction(Nom,Position),
                   color=Nom,linetype=Position))+
  geom_line(position="identity",size=1)+
  geom_point(position="identity",size=3,color="black")+
  geom_point(data=SousDF_LFHF,aes(x=Cdt,y=`LF/HF`,shape=Position),size=2)+
  facet_wrap(vars(Nom))+
  geom_errorbar(aes(x=Cdt,
                    ymin=Moyenne.LFHF-Sd.LFHF,
                    ymax=Moyenne.LFHF+Sd.LFHF),
                size=0.8,width=0.2)+
  labs(title="Moyennes des rapports LF/HF des participants selon la phase et la position",
       subtitle="Barres d'erreur: Sd",
       x="Sujet",
       y=paste("Moyenne ratio LF/HF"),
       caption="WBR,MS,MM")

list_graphs=function(liste_des_objets){
  ls_graphs=list()#vecteur des graphiques
  for (i in liste_des_objets){
    if(is.ggplot(get(i))){
      ls_graphs=append(ls_graphs,i)
      #print(paste(i,"est du type",class(get(i))))
    }
  }
  return(ls_graphs) #retourne une liste de 2 vecteurs, des tabkes(en pos 1), des graphs (en pos 2)
}
Liste_figures=list_graphs(ls())

#fonction sauvegarde des graphiques
graphsave=function(liste_de_graphs,extension){
  for(i in liste_de_graphs){

    print(get(i))
    print(i)
    ggsave(paste("./img/",i,"_",Unique_subject_id,extension,sep=""),
           units=c("in"),width=12,height=8)
    print(paste(i,"sauvegardé dans: ../img"))
  }
}
# 
graphsave(Liste_figures,".png")

