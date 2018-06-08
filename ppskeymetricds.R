#connect to database
source("H:/R setup/ODBC Connection.R")
#call lib
library(tidyverse)

#get data from rowdata no extra
rowdat <- sqlQuery(Dev_Datamart,
                   "select a.year_id,y.year_name,
		                a.row_id, a.auc_id,
                   r.variable, a.value
                    from rowdata_no_extras a 
                    inner join row_ref r 
                   on a.row_id=r.rowid
                   inner join Years y
                   on a.year_id=y.year_ID
                  where  row_id in (42,43,44,45,46,
                     108,
                   37, 153,
                   882,883,
                   376,
                   375,81)
 ")

#get to lead structure
leadstruct <- sqlQuery(BIMSUTST,"select b.ATOMS,b.LEVEL4,b.LEVEL3, d.MANAGE_LEVEL_3,e.LEVEL_3_RC_ADMINISTERING_NAME,d.LEVEL_4_DISPLAY_NAME
                       from ATOMS_UPLOADED b
                         inner join DATAMART_LEVEL4_NEW d
                       on b.LEVEL4=d.LEVEL_4_CODE
                       inner join DATAMART_LEVEL3_NEW e
                       on d.MANAGE_LEVEL_3=e.LEVEL_3_CODE")

library(stringi)
leadstruct <- leadstruct %>%
  mutate(MANAGE_LEVEL_3= stri_pad(as.character(MANAGE_LEVEL_3),2, pad='0'),
         LEVEL3= stri_pad(as.character(LEVEL3),2, pad='0'),
         Dept= paste0(as.character(LEVEL4),'-', as.character(LEVEL_4_DISPLAY_NAME)),
         leadcoll = paste0(MANAGE_LEVEL_3,'-', as.character(LEVEL_3_RC_ADMINISTERING_NAME)))

#by dept
rowdatLead<- merge(rowdat, leadstruct, by.x = 'auc_id', by.y = 'ATOMS')%>%
  group_by(year_id, year_name, row_id, variable, LEVEL4, MANAGE_LEVEL_3, Dept, leadcoll)%>%summarise(value=sum(value))
#by coll
rowdatLead_C<- merge(rowdat, leadstruct, by.x = 'auc_id', by.y = 'ATOMS')%>% mutate(LEVEL4='999',
                                                                                    Dept= paste0(leadcoll,'-All'))%>%
  group_by(year_id, year_name, row_id, variable, LEVEL4, MANAGE_LEVEL_3, Dept, leadcoll)%>%summarise(value=sum(value))

rowdatLead <- rbind(as.data.frame(rowdatLead), as.data.frame(rowdatLead_C))%>%
  arrange(year_id, row_id, MANAGE_LEVEL_3,LEVEL4)


#compute for ratio variables
rowdat1 <- rowdatLead %>%
  #mutate(idgrp = ifelse(row_id %in% c(37,153), 3700, 
   #                     ifelse(row_id %in% c(375,81), 3750, row_id)))%>%
  dcast(year_id+year_name+LEVEL4+ MANAGE_LEVEL_3+ Dept+leadcoll~ row_id, value.var = 'value' )%>%
  mutate(`1080`= ifelse(! is.na(`37`) & `37`>0 & ! is.na(`108`), `108`/`37`,NA ),
          `3700`= ifelse(! is.na(`153`) & `153`>0 & ! is.na(`37`), `37`/`153`,NA ),
         `3750` = ifelse(! is.na(`81`) & `81`>0 & ! is.na(`375`), `375`/`81`,NA ),
         `8820` = ifelse(! is.na(`883`) & `883`>0 & ! is.na(`882`), `882`/`883`,NA ))%>%
  melt(id.vars=c('year_id','year_name','LEVEL4', 'MANAGE_LEVEL_3', 'Dept', 'leadcoll'))%>% filter(! is.na(value))%>%
  #maxyear
  group_by(variable)%>% mutate(maxyr= max(year_id))%>%
  ungroup()%>% filter(maxyr==year_id)

#get variable label
label <- rowdat %>% select(variable, row_id)%>% unique()%>% rename(labelvar= variable, variable= row_id)
rowdat1 <- rowdat1 %>% mutate(variable = as.numeric(as.character(variable)))%>% left_join(label, by='variable')%>%
  mutate(labelvar= ifelse(! is.na(labelvar), as.character(labelvar),
                          ifelse(variable==3700, 'FY Admin Based SCH per Rank Fac',
                                 ifelse(variable==3750, 'Total Grant PI Based/TS AF-FTE', 
                                        ifelse(variable==1080,'General Fund Budget/FY Admin SCH',
                                               ifelse(variable ==8820, 'Tuition Revenue/Instructional Cost',NA))))))%>% 
  filter(! variable %in% c(37,153,375,81,108,882,883)
         # include only instruction college
         & MANAGE_LEVEL_3 %in% c('38', '08','10','02','14','32','16','04','24','46','28','33','34','22','30','06')
         )%>%
  #get short lable
  mutate(shlabelvar= 
           ifelse(labelvar== 'Undergraduate Students', 'Undergraduates',
                  ifelse(labelvar=='Masters Students','Masters',
                         ifelse(labelvar=='Doctoral Students','Doctoral',
                                ifelse(labelvar=='Graduate Professional Students','Grad_Prof',
                                       ifelse(labelvar=='Total Students','Total_Student',
                                              ifelse(labelvar=='Total Grants - 3 Year Average-PI-Real$','V4',
                                                     ifelse(labelvar=='General Fund Budget/FY Admin SCH','V1',
                                                            ifelse(labelvar=='FY Admin Based SCH per Rank Fac','V2',
                                                                   ifelse(labelvar=='Total Grant PI Based/TS AF-FTE','V5','V3'))))))))))





#%>% select(-c(variable))%>%
 # melt(id.vars='labelvar')%>% mutate(var = paste0(labelvar, '_', as.character(variable)),
  #                                   t='t')%>% dcast(t ~ var, value.var = 'value')%>% select(-c(t))

#rowdat2 <- rowdat1 %>% mutate(var= paste0(labelvar,' ', as.character(year_name)))%>%
 # dcast(  MANAGE_LEVEL_3+LEVEL4+ Dept+leadcoll ~ var)

rowdat2a <- rowdat1 %>%
  dcast(  MANAGE_LEVEL_3+LEVEL4+ Dept+leadcoll ~ shlabelvar, value.var = 'value')
rowdat2b <- rowdat1 %>%
  dcast(  MANAGE_LEVEL_3+LEVEL4+ Dept+leadcoll ~ shlabelvar, value.var = 'year_name')
names(rowdat2b)[! names(rowdat2b) %in% c('MANAGE_LEVEL_3','LEVEL4', 'Dept','leadcoll')]<- 
  paste0(names(rowdat2b)[! names(rowdat2b) %in% c('MANAGE_LEVEL_3','LEVEL4', 'Dept','leadcoll')],'_yr')


rowdat2 <- merge(rowdat2a, rowdat2b, by=c('MANAGE_LEVEL_3','LEVEL4', 'Dept','leadcoll'))%>%
  #sorting department order
  mutate(orderdp= ifelse(grepl('All', Dept),1,0))%>%
  arrange(MANAGE_LEVEL_3, orderdp, Dept)

##
#keymetricdat <- rowdat2%>%
  #cbind(rowdat2, qt)%>%
 # rename(Doctoral=`Doctoral Students 2017-18`,
  #       V2=`FY Admin Based SCH per Rank Fac 2016-17`,
   #      V1=`General Fund Budget/FY Admin SCH 2016-17`,
    #     Graduate_Prof=`Graduate Professional Students 2017-18`,
     #    Masters=`Masters Students 2017-18`,
      #   V5=`Total Grant PI Based/TS AF-FTE 2016-17`,
       #  V4=`Total Grants - 3 Year Average-PI-Real$ 2016-17`,
        # Total_Students=`Total Students 2017-18`,
         #V3=`Tuition Revenue/Instructional Cost 2016-17`,
         #Undergrad=`Undergraduate Students 2017-18`
         #V2_1_3=`FY Admin Based SCH per Rank Fac_P1_3`,
         #V2_2_3=`FY Admin Based SCH per Rank Fac_P2_3`,
         #V1_1_3=`General Fund Budget/FY Admin SCH_P1_3`,
         #V1_2_3=`General Fund Budget/FY Admin SCH_P2_3`,
         #V5_1_3=`Total Grant PI Based/TS AF-FTE_P1_3`,
         #V5_2_3=`Total Grant PI Based/TS AF-FTE_P2_3`,
         #V4_1_3=`Total Grants - 3 Year Average-PI-Real$_P1_3`,
         #V4_2_3=`Total Grants - 3 Year Average-PI-Real$_P2_3`,
         #V3_1_3=`Tuition Revenue/Instructional Cost_P1_3`,
         #V3_2_3=`Tuition Revenue/Instructional Cost_P2_3`
         
        # )



##compute quantitle
qt <- rowdat1%>% mutate(advalue = ifelse(LEVEL4=='999',NA, value))%>%
  group_by(labelvar, variable)%>%
  summarise(
    P1_3= quantile(advalue, na.rm = T, prob=1/3),
    P2_3= quantile(advalue, na.rm = T, prob=2/3))%>%
  ungroup()%>% filter(variable %in% c(376,1080,3700,8820,3750))%>%
  mutate(var = ifelse(variable==1080,'V1',
                      ifelse(variable==3700,'V2',
                             ifelse(variable==8820,'V3',
                                    ifelse(variable=='376','V4','V5')))))%>%
  select(-c(labelvar, variable))%>% melt(id.vars='var')%>% arrange(var, variable)%>%
  mutate(variable = substr(as.character(variable),2, nchar(as.character(variable))))%>%
  mutate(vp = paste0(var,'_',variable))




library(xlsx)
write.xlsx2(as.data.frame(rowdat2),
            file="O:/IS/Internal/Reporting/Annual Reports/Key Metrics Set/Data/PPSmetrics.xlsx",sheetName = 'PPS',
            row.names = F)

write.xlsx2(as.data.frame(qt),
            file="O:/IS/Internal/Reporting/Annual Reports/Key Metrics Set/Data/PPSmetrics.xlsx",sheetName = 'PPSquantitle',
            row.names = F, append = T)
