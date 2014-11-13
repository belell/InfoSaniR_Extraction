##########################################################################################
################Extraction des données d'une année de tous les centres sanitaires#########
##########################################################################################
InfoSaniR_extract <- function(directory,health) {
  
  #Folder configuration
  #####################
  #1-Create a main folder called 'InfosaniR'
  #2-Create a sub-folder called 'Script' that contains all R scripts
  #3-Create a sub-folder called 'Data' that contains all Infosani files
  #5-Create a sub-folder called 'Exports' that contains all exported data
  
  
  #########################################FRENCH version####################################
  #I-Description: La fonction extraie automatiquement les données issues du fichier InfoSani# 
  ##############  Excel respectivement les données de pathologie, santé maternité infantile,#
  #               malunitrition et vaccination. Données sont initialement dans un sous      #
  #               du dossier data.                                                          #
  #                                                                                         #
  #II-Arguments: Les arguments d'entrées                                                    #
  #############                                                                             #
  #  II.1-directory: Un caractère qui spécifie la direction du dossier de travail principal #
  #  II.2-health:    Un caractère qui spécifie les données à extraire dans le fichier       #
  #                  Infosani. Ces caractères sont respectivement les suivants:             #
  #     II.2.a-health=='patho' provient de la table 30 (maladies et syndromes sous          #
  #                    surveillance continue) du fichier InfoSani                           #
  #     II.2.b-health=='accouch1' provient de la table 5 (type et lieu accouchement) et la  #
  #                    table 7 (issue de la grossesse) du fichier InfoSani                  #
  #     II.2.c-health=='accouch2' provient de la table 7 (issue de l'accouchement) et la    #
  #                    du fichier InfoSani                                                  #
  #     II.2.d-health=='malnu' provient de la table 16 (Dépistage du suivi de la            #
  #                     malnutrition) du fichier InfoSani                                   #
  #     II.2.e-health=='vac' provient de la table 17 (vaccination) du fichier InfoSani      #     
  #                                                                                         #
  #III-Sorties: Les données extraites sont exportés sous deux formats excel et rdata dans   #                                                                         
  #             un sous dossier appelé 'Exports' du dossier principal.                      #
  ########################################################################################### 
  
  #########################################English version###################################
  #I-Description: The function extracts automatically the data from InfoSani file Excel     # 
  ##############  respectively diseases, childhood maternity health, malnutrition and       #
  #               vaccination data. The data are initially under un folder data             #
  #                                                                                         #
  #                                                                                         #
  #II-Inputs:                                                                               #
  ###########                                                                               #
  #  II.1-directory: A character that specifies the main working directory                  #
  #  II.2-health:    A character that specifies the data to extract in InfoSani file.       #
  #                  These characters are respectively followed                             #
  #     II.2.a-health=='patho' from the table 30 (maladies et syndromes sous                #
  #                    surveillance continue) of InfoSani file                              #
  #     II.2.b-health=='accouch1' from table 5 (type et lieu accouchement) and the table 7  #
  #                    table 7 (issue de la grossesse) of InfoSani file                     #
  #     II.2.c-health=='accouch2' from the table 7 (issue de l'accouchement) and the        #
  #                    of InfoSani file                                                     #
  #     II.2.d-health=='malnu' from the table 16 (Dépistage du suivi de la                  #
  #                     malnutrition) of InfoSani file                                      #
  #     II.2.e-health=='vac' from the table 17 (vaccination) of InfoSani file               #     
  #                                                                                         #
  #III-Ouputs: The extracted data are exported under two formats mainly excel and rdata  in #                                                                         
  #             a subfolder called 'Exports' of main folder                                 #
  ########################################################################################### 
  
  #Chargement des packages
  library(xlsx)
  library(reshape)
  library(stringr)
  
  #Valididation des inputs
  if (!is.character(directory)|!is.character(health)){stop("Directory ou health n'est pas un caractère!")}
  if (!is.element(health,c('patho','accouch1','accouh2','malnu','vac'))){stop("l'argument health mal spécifié!")}
  
  #Extraction de la liste fichiers InfoSani et la date
  FileCenter <- list.files(file.path(getwd(),'Data', directory))
  year <- str_extract(FileCenter, "20[0-9]+") 
  
  #Récupération des départements
  nameDept <- str_sub(FileCenter, start=1,end=str_locate(FileCenter, "_")[,"start"]-1)
  #Définition Département
  nameDept2 <- rep("Abanga-Bine", length(nameDept)) 
  nameDept2[str_detect(nameDept, "OL")] <- "Ogouee-Lacs"
  
  
  #Récupération de type strucutres
  nametypestruc<-c()
  for (i in 1:length(FileCenter)){
    nametypestruc[i]<-str_sub(FileCenter[i], start =str_locate_all(FileCenter, "_")[[i]][1,"start"][[1]]+1,
      end = str_locate_all(FileCenter, "_")[[i]][2,"start"][[1]]-1)}
  #Définition les types d'établissements de santé
  Centertype <- rep("Autre", length(nametypestruc)) 
  Centertype[str_detect(nametypestruc, "Disp")] <- "Dispensaire"
  Centertype[str_detect(nametypestruc, "Inf")] <- "Infirmerie"
  Centertype[str_detect(nametypestruc, "CS")] <- "Centre de Sante"
  Centertype[str_detect(nametypestruc, "CM")] <- "Centre Medical"
  Centertype[str_detect(nametypestruc, "Hop")] <- "Hopital"
  
  #Vérification des fichiers dans le fichier
  if (length(FileCenter)==0){stop("Dossier data vide!")}
  
  #Vérification des fichiers Infosani (formats, structure du nom)
  #Structure du nom:Deptement_TypedeStructure_Nom_Annee.xls
  if (length(FileCenter)!=0){
    ext.fichier<-str_detect(FileCenter,'.xls')
    struc.fichier<-str_count(FileCenter, '_')
    if (sum(ext.fichier)!=length(FileCenter)){stop("Fichier au format non conforme identifié!")}
    else {if (sum(struc.fichier==3|struc.fichier==4)!=length(FileCenter))
                                             {stop("Nom du fichier InfoSani mal structuré!")}}}
  
  #Récupération des noms de structures
  namestruc <- c()
  for (i in 1:length(FileCenter)){
    namestruc[i]<-str_sub(FileCenter[i], start =str_locate_all(FileCenter, "_")[[i]][2,"start"][[1]]+1,
      end = str_locate_all(FileCenter, "_")[[i]][3,"start"][[1]]-1)}
  
  #Définition des soins médicaux (consultations ou hospitalisations)
  soin <- rep("consultation",length(FileCenter))
  soin[str_detect(FileCenter, "Hosp")] <- "hospitalisation"
  
  #Augmentation de la mémoire allouée à R
  options(java.parameters = "-Xmx2048m") 
  ##########################################################################################################
  if (health == 'patho'){
  
  ##Extraction des données de pathologies
  #Extraction des données de tous les établissements de Santé
  ###########################################################
    sheetlist <- c(2,3,4,6,7,8,10,11,12,14,15,16)
    indextab <- data.frame(strow=c(1,54,73,80,126,127,133,159,186,211,224,243,282,301,323,343,355,366,370,371,378,400,435,443))
    indextab$enrow <-c(50,65,76,121,126,132,155,178,207,220,235,278,293,319,339,347,365,369,370,374,392,431,439,449)
    indextab$coldeas<-c(2,2,2,2,2,4,2,2,2,2,2,2,2,2,2,2,2,4,2,4,2,2,2,2)
    
    tab0 <- vector("list",length(FileCenter))
    tab1 <- vector("list",length(sheetlist)*dim(indextab)[1])
    start <- Sys.time()
    
    for (k in 1:length(FileCenter)){
      cat("...Extraction des données:",namestruc[k], "\n")
      
      for (i in 1:12){
        sheet<- read.xlsx(file.path(getwd(),'Data', directory, FileCenter[k]), sheetIndex = sheetlist[i], startRow=556,endRow=1004,
          colIndex=c(1:27),header=FALSE)
        for (j in 1:24){
          tab1[[((i-1)*24+j)]]<-sheet[c(indextab[j,1]:indextab[j,2]),c(1,indextab[j,3],8:27)] 
          names(tab1[[((i-1)*24 +j)]]) <- c("Code_Maladie", "Maladies", "0_CAS_M", "0_CAS_F",
            "0_DCD_M","0_DCD_F", "1_CAS_M", "1_CAS_F", "1_DCD_M","1_DCD_F",
            "5_CAS_M", "5_CAS_F", "5_DCD_M","5_DCD_F", "15_CAS_M", "15_CAS_F",
            "15_DCD_M","15_DCD_F", "50_CAS_M", "50_CAS_F", "50_DCD_M","50_DCD_F")
        }
        rmtm <- round((length(FileCenter)*12 - (k-1)*12 - i)*(difftime(Sys.time(),start,units="mins"))/((k-1)*12 + i),1)
        cat("... Extraction des données du mois: ", month.name[i], "...estimation du temps restant:", rmtm, "minutes \n")      
      }
      tab11<- do.call(rbind, tab1)
      tab11$Code_Domaine<-rep(c("I","N","H","D","V","R","M","S","U","L","E","T","O","C","E'","G","P","Y"),
        c(50,12,4,72,20,22,10,12,36,12,19,17,5,20,15,32,5,7))
      tab11$Domaine<-rep(c("Infect","Metab","Haemato","Gastro","Cardio","Resp","Psych","Neuro","Uro","Ortho","Trauma","Dermato"
        ,"Opthalmo","ORL","Stomato","Gyn","Perinat","Autres"),c(50,12,4,72,20,22,10,12,36,12,19,17,5,20,15,32,5,7))
      tab11$Mois<-rep(month.name,each=370)
      
      tab12 <- melt(tab11, id.vars=c("Code_Maladie","Code_Domaine","Domaine","Maladies","Mois"))
      tab12$Age <- str_extract(tab12$variable, "[0-5]{1,2}")
      tab12$Statut_vital <- str_extract(tab12$variable, "[A-S]{3}")
      tab12$Sexe <- str_extract(tab12$variable, "[M,F]")
      tab12$Annee <- year[k]
      tab12$Nom <- namestruc[k]
      tab12$Departement <- nameDept2[k]
      tab12$Etablissement <- Centertype[k]
      tab12$Soin <- soin[k]
      tab12 <- tab12[,c(13,14,12,15,11,5,2,3,1,4,9,10,8,7)]
      tab0[[k]] <- tab12
      cat("Fin d'extraction des données!\n") 
    }
    tabfinal <- do.call(rbind,tab0)
    tabfinal1 <- tabfinal[!is.na(tabfinal$value),]
    
    #Exportation des donées vers une table Excel/RData
    ##################################################
    tabfinal2 <- tabfinal1
    tabfinal2$Age <- as.factor(tabfinal2$Age)
    tabfinal2$Maladies <- as.character(tabfinal2$Maladies)
    levels(tabfinal2$Age) <- c("0-11 mois","1-4 ans","15-49 ans","5-14 ans",">49 ans")
    names(tabfinal2)[14] <- "Effectif"
    tabfinal3<-tabfinal2
    tabfinal3$Maladies<-iconv(tabfinal3$Maladies,from="UTF-8",to="WINDOWS-1252")#Encodage de la variable 
    write.xlsx2(tabfinal3,file=paste(paste("Exports\\Data_patho",year[1],sep=""),"xlsx",sep="."),row.names=F)
    save(ls="tabfinal2",file=paste(paste("Exports\\Data_patho",year[1],sep=""),"RData",sep="."))
    return(tabfinal2)
  }
  
  ##########################################################################################################
  if (health=='accouch1'){
    
    ##Extraction des données d'accouchement
    #######################################
    data_final<-vector('list',12*length(FileCenter))
    IndexFeuil<-c(2,3,4,6,7,8,10,11,12,14,15,16)
    k<-1  
    
    for(i in 1:length(FileCenter)){cat("...Extraction des données:",namestruc[i],"\n")
      cat("\n")
      j<-1
      while(j<=12){ cat("Données",month.name[j],"\n")
        data<-read.xlsx(file.path(getwd(),'Data',directory,FileCenter[i]),header=F,sheetIndex=IndexFeuil[j],
          colIndex=c(1,8,14,20),rowIndex=c(206:207,211:213,215:224))
        colnames(data)<-c("Type_Accouchement","Domicile","Infrastructure","Autre")
        data_bis<-melt(data,id="Type_Accouchement")
        colnames(data_bis)<-c("Type_Accouchement","Lieu","Effectif")
        data_bis$Annee<-year[i]
        data_bis$Mois<-month.name[j]
        data_bis$Nom<-namestruc[i]
        data_bis$Departement<-nameDept2[i]
        data_bis$Etablissement<-Centertype[i]
        data_final[[k]]<-data_bis[,c(7,8,6,4,5,1:3)]
        k<-k+1; j<-j+1
      } 
      cat("\n")
    }
    print("Tableaux recuperés avec succès!")           
    
    #Exportation des donées vers une table Excel/RData
    ##################################################
    data_merge=do.call('rbind',data_final)    
    last_data<-na.omit(data_merge)
    save(last_data,file=paste(paste("Exports\\Data_Accouch1",year[1],sep=""),"RData",sep="."))
    last_data1<-last_data
    last_data1$Type_Accouchement<-iconv(last_data1$Type_Accouchement,from="UTF-8",to="WINDOWS-1252")
    write.xlsx2(last_data1,file=paste(paste("Exports\\Data_Accouch1",year[1],sep=""),"xlsx",sep="."),row.names=F)
    return(last_data)
    
  } 
  
  ##########################################################################################################
  if (health=='accouch2'){
    
    ##Extraction des données d'accouchement
    #######################################
    data_final<-vector('list',12*length(FileCenter))
    IndexFeuil<-c(2,3,4,6,7,8,10,11,12,14,15,16)
    k<-1  
    
    for(i in 1:length(FileCenter)){cat("...Extraction des données:",namestruc[i],"\n")
      cat("\n")
      j<-1
      while(j<=12){ cat("Données",month.name[j],"\n")
        data<-read.xlsx(file.path(getwd(),'Data',directory,FileCenter[i]),header=F,sheetIndex=IndexFeuil[j],
          colIndex=c(1,9,11,15,17,21,23),rowIndex=c(231:235,237:240))  
        colnames(data)<-c("Naissance","Dom_M","Dom_F","Inf_M","Inf_F","Aut_M","Aut_F")
        data_bis<-melt(data,id="Naissance")
        colnames(data_bis)[3]<-'Effectif'
        data_bis$Lieu <- str_extract(data_bis$variable, "[A-z]{3}")
        data_bis$Lieu[data_bis$Lieu=='Dom']<-'Domicile'
        data_bis$Lieu[data_bis$Lieu=='Inf']<-'Infrastructure'
        data_bis$Lieu[data_bis$Lieu=='Aut']<-'Autre'
        data_bis$Sexe<- str_extract(data_bis$variable, "[M,F]")
        data_bis$Annee<-year[i]
        data_bis$Mois<-month.name[j]
        data_bis$Nom<-namestruc[i]
        data_bis$Departement<-nameDept2[i]
        data_bis$Etablissement<-Centertype[i]
        data_final[[k]]<-data_bis[,c('Departement','Etablissement','Nom','Annee','Mois',
          'Naissance','Lieu','Sexe','Effectif')]
        k<-k+1; j<-j+1     
      }
      cat("\n")
    }
    print("Tableaux recuperés avec succès!") 
    
    #Exportation des donées vers une table Excel/RData
    ##################################################
    data_merge=do.call('rbind',data_final)    
    last_data<-na.omit(data_merge)
    save(last_data,file=paste(paste("Exports\\Data_Accouch2",year[1],sep=""),"RData",sep="."))
    last_data1<-last_data
    last_data1$Naissance<-iconv(last_data1$Naissance,from="UTF-8",to="WINDOWS-1252")
    write.xlsx2(last_data1,file=paste(paste("Exports\\Data_Accouch2",year[1],sep=""),"xlsx",sep="."),row.names=F)
    return(last_data)
  }
  
  ##########################################################################################################
  if (health=='malnu'){
   
    ##Extraction des données malnutrition
    #####################################
    data_final<-vector('list',12*length(FileCenter))
    IndexFeuil<-c(2,3,4,6,7,8,10,11,12,14,15,16)
    k<-1  
    
    for(i in 1:length(FileCenter)){cat("...Extraction des données:",namestruc[i],"\n")
      cat("\n")
      j<-1
      while(j<=12){ cat("Données",month.name[j],"\n")
        data<-read.xlsx(file.path(getwd(),'Data',directory,FileCenter[i]),header=F,sheetIndex=IndexFeuil[j],
          colIndex=c(1,12,17,22),rowIndex=311:317)
        colnames(data)<-c("Depistage","0-11","12-23","24-59")
        data_bis<-melt(data,id="Depistage")
        colnames(data_bis)<-c("Depistage","Age_Mois","Effectif")
        data_bis$Annee<-year[i]
        data_bis$Mois<-month.name[j]
        data_bis$Nom<-namestruc[i]
        data_bis$Departement<-nameDept2[i]
        data_bis$Etablissement<-Centertype[i]
        data_final[[k]]<-data_bis[,c('Departement','Etablissement','Nom','Annee','Mois',
          'Depistage','Age_Mois','Effectif')]
        k<-k+1; j<-j+1  
      }
      cat("\n")
    }
    
    print("Tableaux recuperés avec succès!")
    cat("\n")
    
    #Exportation des donées vers une table Excel/RData
    ##################################################
    data_merge=do.call('rbind',data_final)   
    last_data<-na.omit(data_merge)
    save(last_data,file=paste(paste("Exports\\Data_Depist",year[1],sep=""),"RData",sep="."))
    last_data1<-last_data
    last_data1$Depistage<-iconv(last_data1$Depistage,from="UTF-8",to="WINDOWS-1252")
    write.xlsx2(last_data1,file=paste(paste("Exports\\Data_Depist",year[1],sep=""),"xlsx",sep="."),row.names=F)
    return(last_data)
  } 
  
  ##########################################################################################################
  if (health=='vac'){
     
    ##Extraction des données de vaccination
    #######################################
    data_final<-vector('list',12*length(FileCenter))
    IndexFeuil<-c(2,3,4,6,7,8,10,11,12,14,15,16)
    k<-1  
    
    for(i in 1:length(FileCenter)){cat("...Extraction des données:",namestruc[i],"\n")
      cat("\n")
      j<-1
      while(j<=12){ cat("Données",month.name[j],"\n")
        data<-read.xlsx(file.path(getwd(),'Data',directory,FileCenter[i]),header=F,sheetIndex=IndexFeuil[j],
          colIndex=c(1,8,12,16,20,24),rowIndex=320:347)
        colnames(data)<-c("Antigene","0-11","12-23","24-59","Femmes enceintes","Femmes non enceintes")
        data_bis<-melt(data,id="Antigene")
        colnames(data_bis)<-c("Antigene","Patient","Effectif")
        data_bis$Annee<-year[i]
        data_bis$Mois<-month.name[j]
        data_bis$Nom<-namestruc[i]
        data_bis$Departement<-nameDept2[i]
        data_bis$Etablissement<-Centertype[i]
        data_final[[k]]<-data_bis[,c('Departement','Etablissement','Nom','Annee','Mois',
          'Antigene','Patient','Effectif')]
        k<-k+1; j<-j+1  
      }
      cat("\n")
    }
    
    print("Tableaux recuperés avec succès!")
    cat("\n")
    
    #Exportation des donées vers une table Excel/RData
    ##################################################
    data_merge=do.call('rbind',data_final)    
    last_data<-na.omit(data_merge)
    save(last_data,file=paste(paste("Exports\\Data_Vaccin",year[1],sep=""),"RData",sep="."))
    last_data1<-last_data
    last_data1$Depistage<-iconv(last_data1$Antigene,from="UTF-8",to="WINDOWS-1252")
    write.xlsx2(last_data1,file=paste(paste("Exports\\Data_Vaccin",year[1],sep=""),"xlsx",sep="."),row.names=F)
    return(last_data)
  }
  
}
  
  
  
  
  
