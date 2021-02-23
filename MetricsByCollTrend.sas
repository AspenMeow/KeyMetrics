/***Connect to db***/
%include "C:\Users\chendi4\OneDrive - Michigan State University\From H\SAS\SAS_Log_EDW.sas" ;
LIBNAME BIMSUTST  odbc required="DSN=BIMSUTST; uid=&BIMSUTST_uid; pwd=&BIMSUTST_pwd";
libname PAG oracle path="MSUEDW" user="&MSUEDW_uid" pw= "&MSUEDW_pwd"
schema = OPB_PERS_FALL preserve_tab_names = yes connection=sharedread;
libname Devdm odbc required="DSN=NT_Dev_Datamart";
libname PPS odbc required="DSN=NT_PPS";
libname NonAgr odbc required="DSN=NT_Non_Aggregated";
libname Cwalk odbc required="DSN=NT_Crosswalks";

%let todaysDate = %sysfunc(today(), yymmddn8.);

option symbolgen;
options mlogic;


/*set pag year parameter value
%let entrycohort =2012;
%let gradcohort ='2018-2019';*/
/*set up pag yr paramter dynamically*/
proc sql stimer;
select max(cohort) into: entrycohort
from PAG.PERSISTENCE_V
 where (ENTRANT_SUMMER_FALL='Y' or substr(ENTRY_TERM_CODE,1,1)='F') 
 and student_level='UN' and level_entry_status='FRST' and grad6 is not null;

 select distinct PARAMETER_VALUE_TEXT into: gradcohort
 from PAG.STANDARD_REPORT_PARAMETERS_V
 where parameter_name ='PAG_GRADUATING_COHORT_CEILING_YEAR';

 /*select distinct PARAMETER_VALUE_TEXT into: AAver
 from PAG.STANDARD_REPORT_PARAMETERS_V
 where parameter_name ='ACAD_ANALYTIC_AAU_COMP_VERSION';*/
 quit;


/***********get data from rowdata no extra*******************************/
proc sql stimer;
 create table ppsraw as 
 select a.year_id,y.year_name,
		                a.row_id, a.auc_id,
                   r.variable, a.value, at.LEVEL4,d.MANAGE_LEVEL_3,e.LEVEL_3_RC_ADMINISTERING_NAME,d.LEVEL_4_DISPLAY_NAME
                    from Devdm.rowdata_no_extras a 
                    inner join Devdm.row_ref r 
                   on a.row_id=r.rowid
                   inner join Devdm.Years y
                   on a.year_id=y.year_ID
				   inner join BIMSUTST.ATOMS_UPLOADED at 
				   on a.auc_id=at.ATOMS
				   inner join BIMSUTST.DATAMART_LEVEL4_NEW d
                    on at.LEVEL4=d.LEVEL_4_CODE
                   inner join BIMSUTST.DATAMART_LEVEL3_NEW e
                       on d.MANAGE_LEVEL_3=e.LEVEL_3_CODE
                  where  row_id in (42,43,44,45,46,
                   108,
                   37, 153,
                   882,883,
                   376,
                   375,81,
				   /* new faculty cnt*/
5,6,487,488,19);

/*by college and dept*/
   create table ppsdp as 
   select distinct year_id, year_name, row_id, variable, trim(MANAGE_LEVEL_3)||'-'|| trim(LEVEL_3_RC_ADMINISTERING_NAME) as leadcoll, trim(LEVEL4)||'-'|| trim(LEVEL_4_DISPLAY_NAME) as Dept, sum(value) as value
   from ppsraw
   group by year_id, year_name, row_id, variable, MANAGE_LEVEL_3, LEVEL_3_RC_ADMINISTERING_NAME, LEVEL4, LEVEL_4_DISPLAY_NAME;

/*by college only*/
   create table ppscoll as 
   select distinct year_id, year_name, row_id, variable,trim(MANAGE_LEVEL_3)||'-'|| trim(LEVEL_3_RC_ADMINISTERING_NAME) as leadcoll,  trim(MANAGE_LEVEL_3)||'-'|| trim(LEVEL_3_RC_ADMINISTERING_NAME)||'-All' as Dept ,sum(value) as value
   from ppsraw
   group by year_id, year_name, row_id, variable, MANAGE_LEVEL_3,LEVEL_3_RC_ADMINISTERING_NAME;


/*Univ*/
   create table ppsall as 
   select distinct year_id, year_name, row_id, variable,'University Total' as leadcoll ,'University Total' as dept ,sum(value) as value
   from ppsraw
   group by year_id, year_name, row_id, variable;

quit;




data pps;
set  ppsdp ppscoll ppsall;
run;

/*transpose and compute for ratio**/
proc sort data= pps; by year_id year_name leadcoll dept; run;
proc transpose data=pps out=ppsl prefix=V;
 by year_id year_name leadcoll dept;
 id row_id;
 var value;
 run;
data ppsl;
set ppsl;
if V108=. or V37=. or V37=0 then V1080=. ;
else V1080=V108/V37;
if V37=. or V153=. or V153=0 then V3700=.;
else V3700=V37/V153;
if V375=. or V81=. or V81=0 then V3750=.;
else V3750=V375/V81;
if V882=. or V883=. or V883=0 then V8820=.;
else V8820=V882/V883;
V4878= V487 +V488;
drop _NAME_;
run;
/**transpose back to long**/
proc transpose data=ppsl out =pps;
by year_id year_name leadcoll dept;
run;
data pps;
set pps;
/**filter null value and numerator deminoaor */
where COL1 ne . and _NAME_ not in ('V37','V153','V375','V81','V108','V882','V883','V487','V488');
run;
/**get max year**/
proc sql stimer;
create table yr as 
select a.*,b.rowid as row_id
from PPS.DataAllYears a
inner join Devdm.row_ref b
on a.typea=b.typea
and a.subtype=b.subtype
and a.type=b.type
and a.suborder=b.suborder
where a.infounit='51000'
and b.rowid in (42,43,44,45,46,
                   108,
                   37, 153,
                   882,883,
                   376,
                   375,81,
5,6,487,488,19);
quit;

proc contents data=yr out=var noprint;
run;

proc sql noprint;
select NAME into: yearlist separated by ' '
from var
where substr(NAME,1,4)='Year';
quit;
/*transpose*/
proc sort data=yr; by row_id; run;
proc transpose data= yr out=yrlong;
by row_id;
var &yearlist.;
run;

data yrlong;
set yrlong;
yrid = input( substr(_LABEL_,5, length(_LABEL_)-4),2. );
where col1 ne . ;
run;

proc sort data=yrlong; by row_id yrid; run;
data yrlong;
set yrlong;
by row_id yrid;
if last.row_id then output;
run;

proc transpose data=yrlong out=yrlong1 prefix=V;
id row_id;
var yrid;
run;

data yrlong1;
set yrlong1;
V1080= min(V108,V37);
V3700=min(V37,V153);
 V3750 = min(V375,V81);
V8820=min(V882,V883);
V4878= min(V487 ,V488);
drop _NAME_;
run;

/**transpose back to long**/
proc transpose data=yrlong1 out =yrlong1;
run;


proc sql;
create table pps as 
select a.*,  ( b.COL1-a.year_id ) as yrdiff
from pps a
inner join yrlong1 b
on a._NAME_=b._NAME_ and a.year_id>=b.COL1-5;
quit;

/*proc sort data=pps; by leadcoll Dept _NAME_ year_id;
run;
data pps;
set pps;
by leadcoll Dept _NAME_ year_id ;
if last._NAME_ then output;
run;*/
/**recode variable**/
data pps;
set pps;
if _NAME_='V42' then var='Undergraduates';
else if _NAME_='V43' then var='Masters';
else if _NAME_='V44' then var='Doctoral';
else if _NAME_='V45' then var='Grad_Prof';
else if _NAME_='V46' then var='Total_Student';
else if _NAME_='V1080' then var='V1';
else if _NAME_='V3700' then var='V2';
else if _NAME_='V8820' then var='V3';
else if _NAME_='V376' then var='V4';
else if _NAME_='V3750' then var='V5';

/**new faculty metric**/
else if _NAME_='V5' then var='V6'; /*tenure*/
else if _NAME_='V6' then var='V7'; /*fix term fac*/
else if _NAME_='V4878' then var='V8'; /*other academics fac*/
else if _NAME_='V19' then var='V9'; /*tenure system new hire*/

else var=_NAME_;
/*include only instruction college*/
where substr(leadcoll,1,2) in ('38', '08','10','02','14','32','16','04','24','46','28','33','34','22','30','06') or leadcoll='University Total';
run;


 /*get PPS version macro var*/
proc sql ;
select VERDATE
into :version
from BIMSUTST.version;
quit;

proc sql;
create table varlist as 
select distinct var
from pps;
quit;

data varlist;
set varlist;
length varlabel $35.;
if var = 'Undergraduates' then varlabel = "Undergraduates Fall Enrollment";
else if var = 'Masters' then varlabel = "Masters Fall Enrollment";
else if var = 'Doctoral' then varlabel = "Doctoral Fall Enrollment";
else if var = 'Total_Student' then varlabel = "Total Fall Enrollment";
else if var = 'V1' then varlabel = "General Fund Budget/FY Admin SCH";
else if var = 'V2' then varlabel = "Admin Based SCH per Rank Faculty FY";
else if var = 'V3' then varlabel = "Tuition Revenue/Instructional Cost";
else if var = 'V4' then varlabel = "Total Grants - 3 Year Average-PI-Real$";
else if var = 'V5' then varlabel = "Total Grant PI Based/TS AF-FTE";

else if var = 'V6' then varlabel = "Tenure System Faculty Headcount";
else if var = 'V7' then varlabel = "Fixed Term Faculty Headcount";
else if var = 'V8' then varlabel = "Other Academics Headcount";
else if var = 'V9' then varlabel = "Tenure System New Hire Headcount";


if var = 'Undergraduates' then varseq = 2;
else if var = 'Masters' then varseq = 3;
else if var = 'Doctoral' then varseq = 4;
else if var = 'Total_Student' then varseq = 1;
else if var = 'V1' then varseq = 5;
else if var = 'V2' then varseq = 6;
else if var = 'V3' then varseq = 7;
else if var = 'V4' then varseq = 8;
else if var = 'V5' then varseq = 9;
else if var = 'V6' then varseq = 10;
else if var = 'V7' then varseq = 11;
else if var = 'V8' then varseq = 12;
else if var = 'V9' then varseq = 13;
where var ne 'Grad_Prof';
run;

proc sort data=varlist ; by varseq; run;

/*macro for loop over all var*/
proc sql noprint;
select distinct count(distinct var)
 into :varcnt
 from varlist;
quit;


data _null_;
set varlist;
call symput('var'||left(_n_),var);
call symput('varlabel'||left(_n_),varlabel);
run;

%put &varlabel1;


/**PPS color code**/
data ppscolor;
set pps;
where leadcoll='University Total';
run;

proc sort data= ppscolor; by var yrdiff; run;
proc transpose data=ppscolor out=ppscolor prefix=YR ;
 by var;
 id yrdiff;
 var col1;
 run;
data ppscolor;
set ppscolor;
pct1 = (YR0-YR1)/YR1;
pct5= (YR0-YR5)/YR5;
keep var pct1 pct5;
run;
proc sql;
create table ppscolor as
select a.*, varseq
from ppscolor a
inner join varlist b
on a.var=b.var
order by varseq;
quit;

data _null_;
set ppscolor;
call symput('ppspctoneyr'||left(_n_),pct1);
call symput('ppspctfiveyr'||left(_n_),pct5);
run;
%put &ppspctfiveyr1;


/************************************************/
/****************PAG*******************************/



proc sql stimer;

create table PAG_enterdp as 
select COHORT, strip(COLLEGE_FIRST)||'-'|| strip(COLLEGE_FIRST_NAME) as Coll_1st, strip(DEPT_FIRST) ||'-'|| strip(DEPT_FIRST_NAME) as Dept_1st,
       count(PID) as N, avg(PERSIST1) as PERSIST1,avg(PERSIST2) as PERSIST2,
	   avg(PERSIST3) as PERSIST3, avg(GRAD4) as GRAD4,
	   avg(GRAD5)as GRAD5, avg(GRAD6) as GRAD6
from PAG.PERSISTENCE_V
where Student_LEVEL='UN'
and (  ENTRANT_SUMMER_FALL='Y' or substr(ENTRY_TERM_CODE,1,1)='F' )
and LEVEL_ENTRY_STATUS='FRST'
and cohort >= &entrycohort-5 
group by COHORT, COLLEGE_FIRST, COLLEGE_FIRST_NAME, DEPT_FIRST, DEPT_FIRST_NAME;

create table PAG_entercoll as 
select COHORT, strip(COLLEGE_FIRST)||'-'|| strip(COLLEGE_FIRST_NAME) as Coll_1st, strip(COLLEGE_FIRST)||'-'|| strip(COLLEGE_FIRST_NAME)||'-All' as Dept_1st,
       count(PID) as N, avg(PERSIST1) as PERSIST1,avg(PERSIST2) as PERSIST2,
	   avg(PERSIST3) as PERSIST3, avg(GRAD4) as GRAD4,
	   avg(GRAD5)as GRAD5, avg(GRAD6) as GRAD6
from PAG.PERSISTENCE_V
where Student_LEVEL='UN'
and (  ENTRANT_SUMMER_FALL='Y' or substr(ENTRY_TERM_CODE,1,1)='F' )
and LEVEL_ENTRY_STATUS='FRST'
and cohort >= &entrycohort-5 
group by COHORT, COLLEGE_FIRST, COLLEGE_FIRST_NAME;

create table PAG_enterall as 
select COHORT, 'University Total' as Coll_1st, 'University Total' as Dept_1st,
       count(PID) as N, avg(PERSIST1) as PERSIST1,avg(PERSIST2) as PERSIST2,
	   avg(PERSIST3) as PERSIST3, avg(GRAD4) as GRAD4,
	   avg(GRAD5)as GRAD5, avg(GRAD6) as GRAD6
from PAG.PERSISTENCE_V
where Student_LEVEL='UN'
and (  ENTRANT_SUMMER_FALL='Y' or substr(ENTRY_TERM_CODE,1,1)='F' )
and LEVEL_ENTRY_STATUS='FRST'
and cohort >= &entrycohort-5 
group by COHORT;


create table PAG_Degrdp as 
select GRADUATING_COHORT, strip(COLLEGE_DEGREE)||'-'|| strip(COLLEGE_DEGREE_NAME) as Coll_DEGR,  strip(DEPT_DEGREE) ||'-'|| strip(DEPT_DEGREE_NAME) as Dept_DEGR,
       count(PID) as N,AVG(TTD_IN_YEARS) AS TTD
from PAG.PERSISTENCE_V
where Student_LEVEL='UN'
and (  ENTRANT_SUMMER_FALL='Y' or substr(ENTRY_TERM_CODE,1,1)='F' )
and LEVEL_ENTRY_STATUS='FRST'
and input(substr(Graduating_Cohort,1,4),4.)>= %eval(%substr(&gradcohort,1,4)-5) 
and input(substr(Graduating_Cohort,1,4),4.)<= %substr(&gradcohort,1,4)
group by GRADUATING_COHORT, COLLEGE_DEGREE, COLLEGE_DEGREE_NAME, DEPT_DEGREE, DEPT_DEGREE_NAME;


create table PAG_Degrcoll as 
select GRADUATING_COHORT, strip(COLLEGE_DEGREE)||'-'|| strip(COLLEGE_DEGREE_NAME) as Coll_DEGR,  strip(COLLEGE_DEGREE)||'-'|| strip(COLLEGE_DEGREE_NAME)||'-All' as Dept_DEGR,
       count(PID) as N,AVG(TTD_IN_YEARS) AS TTD
from PAG.PERSISTENCE_V
where Student_LEVEL='UN'
and (  ENTRANT_SUMMER_FALL='Y' or substr(ENTRY_TERM_CODE,1,1)='F' )
and LEVEL_ENTRY_STATUS='FRST'
and input(substr(Graduating_Cohort,1,4),4.)>= %eval(%substr(&gradcohort,1,4)-5) 
and input(substr(Graduating_Cohort,1,4),4.)<= %substr(&gradcohort,1,4)
group by GRADUATING_COHORT, COLLEGE_DEGREE, COLLEGE_DEGREE_NAME;

create table PAG_Degrall as 
select GRADUATING_COHORT, 'University Total' as Coll_DEGR, 'University Total' as Dept_DEGR,
       count(PID) as N,AVG(TTD_IN_YEARS) AS TTD
from PAG.PERSISTENCE_V
where Student_LEVEL='UN'
and (  ENTRANT_SUMMER_FALL='Y' or substr(ENTRY_TERM_CODE,1,1)='F' )
and LEVEL_ENTRY_STATUS='FRST'
and input(substr(Graduating_Cohort,1,4),4.)>= %eval(%substr(&gradcohort,1,4)-5) 
and input(substr(Graduating_Cohort,1,4),4.)<= %substr(&gradcohort,1,4)
group by GRADUATING_COHORT;
run;


/*combine ds*/
data PAG_enter;
set PAG_enterdp  PAG_entercoll PAG_enterall;
run;
 
data PAG_Degr;
set PAG_DegrDp PAG_Degrcoll PAG_Degrall;
run;

/*transpse*/
proc sort data = pag_enter; by cohort coll_1st dept_1st;
proc transpose data=PAG_enter out =PAG_enter;
by cohort coll_1st dept_1st;
run;

proc sort data = pag_degr; by graduating_cohort coll_degr dept_degr;
proc transpose data=PAG_degr out =PAG_degr;
by graduating_cohort coll_degr dept_degr ;
run;

data pag_enter;
set pag_enter;
cohort1 = put(cohort, 4.);
run; 
proc sql;
create table  pagset as 
select cohort1, coll_1st, Dept_1st,_NAME_, col1
from pag_enter 
where COL1 ne .
union
select *
from pag_degr
where COL1 ne .;
run;

proc sql;
create table pagset as 
select a.*
from pagset a 
inner join (

select  coll_1st, _NAME_, input(max(cohort1),4.)-5 as yrcutoff
from pagset
group by coll_1st, _NAME_ ) b
on a.coll_1st=b.coll_1st and a._NAME_=b._NAME_
where input(cohort1,4.)>= yrcutoff;
quit;

data pagset;
set pagset;
if _NAME_ ne 'TTD' then col1= col1/100;
run;



proc sql;
create table paglist as 
select distinct _NAME_
from pagset;
quit;

data paglist;
set paglist;
if _NAME_ = 'PERSIST1' then do;
varseq =1;
varlabel = "Persist 1st Returning Fall";
end;
if _NAME_ = 'PERSIST2' then do;
varseq =2;
varlabel = "Persist 2nd Returning Fall";
end;
if _NAME_ = 'PERSIST3' then do;
varseq =3;
varlabel = "Persist 3rd Returning Fall";
end;
if _NAME_ = 'GRAD4' then do;
varseq =4;
varlabel = "Graduated by 4th Yr";
end;
if _NAME_ = 'GRAD5' then do;
varseq =5;
varlabel = "Graduated by 5th Yr";
end;
if _NAME_ = 'GRAD6' then do;
varseq =6;
varlabel = "Graduated by 6th Yr";
end;
if _NAME_ = 'TTD' then do;
varseq =7;
varlabel = "Time-to-Degree in Year";
end;
where _NAME_ ne 'N';
run;

proc sort data=paglist ; by varseq; run;

/*macro for loop over all var*/
proc sql noprint;
select distinct count(distinct _NAME_)
 into :pagvarcnt
 from paglist;
quit;

data _null_;
set paglist;
call symput('pagvar'||left(_n_),_NAME_);
call symput('paglabel'||left(_n_),varlabel);
run;

%put &pagvar1;

/**Pag color code**/
data pagcolor;
set pagset;
where coll_1st='University Total';
run;

proc sort data= pagcolor; by _NAME_ cohort1; run;
data pagcolor;
set pagcolor;
count + 1;
 by _NAME_;
 if first._NAME_ then count = 0;
 run;
data pagcolor; set pagcolor; yrdiff= 5- count; run;

proc transpose data=pagcolor out=pagcolor prefix=YR ;
 by _NAME_;
 id yrdiff;
 var col1;
 run;


data pagcolor;
set pagcolor;
pct1 = (YR0-YR1)/YR1;
pct5= (YR0-YR5)/YR5;
keep _NAME_ pct1 pct5;
where _NAME_ ne 'N';
run;
proc sql;
create table pagcolor as
select a.*, varseq
from pagcolor a
inner join paglist b
on a._NAME_=b._NAME_
order by varseq;
quit;

data _null_;
set pagcolor;
call symput('pagpctoneyr'||left(_n_),pct1);
call symput('pagpctfiveyr'||left(_n_),pct5);
run;
%put &pagpctfiveyr1;

/**AA**/
proc sql stimer;
/*AAprogram level data*/
create table AAprog as
select *
from BIMSUTST.ACADEMIC_ANALYTICS_METRICS_DP;

/*for adding college name*/
create table LEVEL3 as 
select *
from BIMSUTST.DATAMART_LEVEL3_NEW;
quit;

data AAprog;
 set AAprog;
 if COLLEGE='AG' then MAU='02';
 else if COLLEGE='BUS' then MAU='08';
 else if COLLEGE='CAL' then MAU='04';
 else if COLLEGE='CHM' then MAU='22';
 else if COLLEGE='CNS' then MAU='32';
 else if COLLEGE='COM' then MAU='34';
 else if COLLEGE='COMM' then MAU='10';
 else if COLLEGE='ED' then MAU='14';
 else if COLLEGE='ENG' then MAU='16';
 else if COLLEGE='MUS' then MAU='30';
 else if COLLEGE='NUR' then MAU='33';
 else if COLLEGE='SOC' then MAU='38';
 else if COLLEGE='VET' then MAU='46';
 /*convert to numeric*/
 PCT_SCHLR_RSCH_INDEX1= input(PCT_SCHLR_RSCH_INDEX, 10.);
 RANK_SCHLR_RSCH_INDEX1= input(RANK_SCHLR_RSCH_INDEX,10.);
 AAU_PROG_IN_DISCIPLINE1= input(AAU_PROG_IN_DISCIPLINE,10.);
 PCT_CITATIONS_PER_PUBLICATION1= input(PCT_CITATIONS_PER_PUBLICATION,10.);
 PCT_JOURNAL_PUBS_PER_FACULTY1= input(PCT_JOURNAL_PUBS_PER_FACULTY,10.);
 PCT_CITATIONS_PER_FACULTY1=input(PCT_CITATIONS_PER_FACULTY,10.);
 NATIONAL_ACADEMY_MEMBERS1=input(NATIONAL_ACADEMY_MEMBERS,10.);
 PCT_GRANTS_PER_FACULTY1=input(PCT_GRANTS_PER_FACULTY,10.);
 PCT_GRANT_DOLLARS_PER_FACULTY1=input(PCT_GRANT_DOLLARS_PER_FACULTY,10.);
 PCT_DOLLARS_PER_GRANT1= input(PCT_DOLLARS_PER_GRANT,10.);
 PCT_AWARDS_PER_FACULTY1= input(PCT_AWARDS_PER_FACULTY,10.);
 PCT_CONF_PROCEED_PER_FACULTY1= input(PCT_CONF_PROCEED_PER_FACULTY,10.);
 PCT_BOOKS_PER_FACULTY1=input(PCT_BOOKS_PER_FACULTY,10.);
 if strip(PROG_IN_MULT_TAXONOMIES) = 'Y' then  PROG_IN_MULT_TAXONOMIES=PROG_IN_MULT_TAXONOMIES;
 else PROG_IN_MULT_TAXONOMIES='N';
run;
/*merge to get college name*/
proc sql stimer;
create table AAprog as 
select a.*,b.LEVEL_3_RC_ADMINISTERING_NAME as MAUName
from AAprog a 
left join Level3 b
on a.MAU=b.LEVEL_3_CODE
order by MAU, Dept;
quit;



proc template;
 define style Styles.Custom;
 parent = Styles.Printer;
 Style SystemTitle from SystemTitle /
Font = ("Arial, Helvetica", 7pt, bold );
Style SystemFooter from systemFooter /
Font = ("Arial, Helvetica", 7pt, bold );
 
class color_list /
'link' = blue /* links */
'bgH' = white /* row and column header background */
'bgT' = white /* table background */
'bgD' = white /* data cell background */
'fg' = black /* text color */
'bg' = white; 
class Table from Output /
frame = hsides /* outside borders: void, box, above/below, vsides/hsides, lhs/rhs */
rules = all /* internal borders: none, all, cols, rows, groups */
cellpadding = 4pt /* the space between table cell contents and the cell border */
cellspacing = 0.2pt /* the space between table cells, allows background to show */
borderwidth = 0.05pt /* the width of the borders and rules */
bordercolor = #D3D3D3
background = color_list('bgT') /* table background color */;
 end;
run;



data pps;
set pps;
dummy ='PPS';
run;

data pagset;
set pagset;
dummy ='Pag';
run;
%macro ppsmetric;
%do k=1 %to &varcnt;
%if &k<10 %then %do;
ods proclabel="PPS Metrics"  ;
%end;
%else %do;
ods proclabel="Faculty and Academic Staff Metrics"  ;
%end;

proc report data=pps style(report)={outputwidth=95% font_size=8pt} style(header)={font_face='Arial, Helvetica' font_size=8pt} contents="&&varlabel&k" ;
%if &k<10 %then %do;
title1 h=13pt "PPS Metrics - &&varlabel&k"  ;
%end;
%else %do;
title1 h=13pt "Faculty and Academic Staff - &&varlabel&k"  ;
%end;
footnote3 h=9pt justify=left "PPS version: &version";
footnote4 j = r 'Page !{thispage} of !{lastpage}';
column dummy leadcoll COL1,year_name ('Percent Change' Fiveyr Oneyr);
define dummy/group noprint;
define leadcoll/group 'Lead MAU' ;
define year_name/'Year' across;
%if &k= 5 %then %do;
define COL1/' ' analysis sum format=dollar15.0 ;
%end;
%else %if &k= 7 %then %do;
define COL1/' ' analysis sum format=percentn15.0 ;
%end;
%else %do;
define COL1/' ' analysis sum format=comma15.0 ;
%end;
define fiveyr/'Five Year' computed format=percentn15.1;
define Oneyr/'One Year' computed format=percentn15.1;
compute fiveyr;
fiveyr = (_c8_- _c3_)/_c3_;
%if &k= 5 %then %do;
if fiveyr lt &&ppspctfiveyr&k & fiveyr ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if fiveyr gt &&ppspctfiveyr&k & fiveyr ne . then call define(_col_,"style","style={background=#fff7bc}");
	 if _c2_="University Total" then  CALL DEFINE("_c9_", "style","STYLE=[BACKGROUND=white]");
%end;
%else %do;
if fiveyr gt &&ppspctfiveyr&k & fiveyr ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if fiveyr lt &&ppspctfiveyr&k & fiveyr ne . then call define(_col_,"style","style={background=#fff7bc}");
	 if _c2_="University Total" then  CALL DEFINE("_c9_", "style","STYLE=[BACKGROUND=white]");
%end;
endcomp;
compute oneyr;
oneyr = (_c8_- _c7_)/_c7_;
%if &k= 5 %then %do;
if oneyr lt &&ppspctoneyr&k & oneyr ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if oneyr gt &&ppspctoneyr&k & oneyr ne . then call define(_col_,"style","style={background=#fff7bc}");
	 if _c2_="University Total" then  CALL DEFINE("_c10_", "style","STYLE=[BACKGROUND=white]");

%end;
%else %do;
if oneyr gt &&ppspctoneyr&k & oneyr ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if oneyr lt &&ppspctoneyr&k & oneyr ne . then call define(_col_,"style","style={background=#fff7bc}");
	 if _c2_="University Total" then  CALL DEFINE("_c10_", "style","STYLE=[BACKGROUND=white]");
%end;
endcomp;
compute leadcoll;
     if leadcoll='University Total' then call define(_row_,"style","style={font_weight=bold}");
	
 endcomp;
 
break before dummy/ page contents=""; /*remove 3rd node*/
where var="&&var&k" and (find(Dept,'All','i') or Dept ='University Total');
run;
%end;

%mend ppsmetric;



%macro pagmetric;
%do k=1 %to &pagvarcnt;
footnote3 h=9pt justify=left "PAG version: SISFull(FS20)";
footnote4 j = r 'Page !{thispage} of !{lastpage}';
ods proclabel="PAG Metrics"  ;

proc report data=pagset  style(report)={outputwidth=95% font_size=8pt} style(header)={font_face='Arial, Helvetica' font_size=8pt} contents="&&paglabel&k" ;
title1 h=13pt "PAG Metrics - &&paglabel&k"  ;

column dummy Coll_1st COL1,cohort1 ('Percent Change' Fiveyr Oneyr);
define dummy/group noprint;
%if &k= 7 %then %do;
define Coll_1st/group 'Degree College' ;
%end;
%else %do;
define Coll_1st/group 'Entering College' ;
%end;
%if &k= 7 %then %do;
define cohort1/'Graduating Cohort' across;
%end;
%else %do;
define cohort1/'Entering Cohort' across;
%end;
%if &k= 7 %then %do;
define COL1/' ' analysis sum format=comma15.2;
%end;
%else %do;
define COL1/' ' analysis sum format=percentn15.1 ;
%end;
define fiveyr/'Five Year' computed format=percentn15.1;
define Oneyr/'One Year' computed format=percentn15.1;

compute fiveyr;
fiveyr = (_c8_- _c3_)/_c3_;
%if &k=7 %then %do;
if fiveyr lt &&pagpctfiveyr&k & fiveyr ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if fiveyr gt &&pagpctfiveyr&k & fiveyr ne . then call define(_col_,"style","style={background=#fff7bc}");
	 if _c2_="University Total" then  CALL DEFINE("_c9_", "style","STYLE=[BACKGROUND=white]");

endcomp;
%end;
%else %do;
if fiveyr gt &&pagpctfiveyr&k & fiveyr ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if fiveyr lt &&pagpctfiveyr&k & fiveyr ne . then call define(_col_,"style","style={background=#fff7bc}");
	 if _c2_="University Total" then  CALL DEFINE("_c9_", "style","STYLE=[BACKGROUND=white]");

endcomp;
%end;
compute oneyr;
oneyr = (_c8_- _c7_)/_c7_;
%if &k=7 %then %do;
if oneyr lt &&pagpctoneyr&k & oneyr ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if oneyr gt &&pagpctoneyr&k & oneyr ne . then call define(_col_,"style","style={background=#fff7bc}");
	 if _c2_="University Total" then  CALL DEFINE("_c10_", "style","STYLE=[BACKGROUND=white]");

%end;
%else %do;
if oneyr gt &&pagpctoneyr&k & oneyr ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if oneyr lt &&pagpctoneyr&k & oneyr ne . then call define(_col_,"style","style={background=#fff7bc}");
	 if _c2_="University Total" then  CALL DEFINE("_c10_", "style","STYLE=[BACKGROUND=white]");
%end;
endcomp;

 compute COLL_1st;
     if Coll_1st='University Total' then call define(_row_,"style","style={font_weight=bold}");
	
 endcomp;
break before dummy/ page contents=""; /*remove 3rd node*/

where _NAME_="&&pagvar&k" and ( find(Dept_1st,'All','i') or Dept_1st ='University Total');
run;
%end;
%mend pagmetric;




ods listing close;
/**output prep for proc document*/
ods document name=mydocrep(write);
%ppsmetric;
%pagmetric;
ods document close;


proc document name=mydocrep ;

run;
options   nodate  nonumber leftmargin=0.5in rightmargin=0.5in orientation=landscape papersize=A4 center missing=' ' ;

ods pdf file="O:\IS\Internal\Reporting\Annual Reports\Key Metrics Set\Output\KeyMetrics_College_Trend&todaysDate..pdf" style=Custom contents=yes;
ods escapechar= '!';
 replay;
run;
ods listing;
ods pdf close;
quit; 



/********************for race ethnicity one year data******************/
proc sql ;
create table hr20 as
SELECT *,(case when RES_STATUS_CD in ('C','N','T') then 'Citizen or Permanent Resident' 
                                            when RES_STATUS_CD in ('A','S') then 'Non-Resident Alien' 
                                            else 'Unknown' end) as Citizen, (case when ETHNC_CD='E1' then ETHNC_NM 
                                          when RACIAL_CAT_CD_2 is not null then 'Two or More Races' 
                                          when RACIAL_CAT_CD_1 is not null then RACIAL_CAT_NM_1 
                                          else 'Unknown' end) as RACE_ETHNICITY, trim(mau_code)||trim(dept_code)||'00' as cuc
  FROM NonAgr.HR_Person_Position_Job
  where CYEAR in ('2020','2015')
  and anl_sal>0
  and emp_sgrp_cd not in ('AT','AS','B4','AC','A1','AW','AD','AX','B7','B8')
and FacStafPosFlag=1
and TempOnCall='N'
and LTD ='N';
quit;

proc sql;
create table hr20cw as 
select a.*,b.Aggregated_Unit_Code as auc_id
from hr20 a
inner join cwalk.Crosswalk_Current b
on a.cuc=b.detail_text;

create table hr20cw as
select a.*, trim(MANAGE_LEVEL_3)||'-'|| trim(LEVEL_3_RC_ADMINISTERING_NAME) as leadcoll, trim(LEVEL4)||'-'|| trim(LEVEL_4_DISPLAY_NAME) as Dept
from hr20cw a
 inner join BIMSUTST.ATOMS_UPLOADED at 
				   on a.auc_id=at.ATOMS
				   inner join BIMSUTST.DATAMART_LEVEL4_NEW d
                    on at.LEVEL4=d.LEVEL_4_CODE
                   inner join BIMSUTST.DATAMART_LEVEL3_NEW e
                       on d.MANAGE_LEVEL_3=e.LEVEL_3_CODE;
quit;

data hr20cw;
set hr20cw;
if emp_cat_cd_1='S' then empgrp=4;
else if Employee_category = 'Fix Staff' or Employee_category ='Cont Staff' then empgrp =3;
else if  Employee_category ='Tenure Sys' then empgrp=1;
else empgrp = 2;
dummy='hr';
Headcount=1;
run;


proc sql;
create table hr20cwagg15 as 
select leadcoll, dept,empgrp, race_ethnicity, cyear,sum(headcount) as headcount15
from hr20cw
where cyear='2015'
group by leadcoll, dept,empgrp, race_ethnicity, cyear
order by leadcoll, dept,empgrp, race_ethnicity, cyear;

create table hr20cwagg20 as 
select leadcoll, dept,empgrp, race_ethnicity, cyear,sum(headcount) as headcount20
from hr20cw
where cyear='2020'
group by leadcoll, dept,empgrp, race_ethnicity, cyear
order by leadcoll, dept,empgrp, race_ethnicity, cyear;

create table hrw as
select (case when a.leadcoll is null then b.leadcoll else a.leadcoll end) as leadcoll,
(case when a.dept is null then b.dept else a.dept end) as dept,
(case when a.empgrp is null then b.empgrp else a.empgrp end) as empgrp,
(case when a.race_ethnicity is null then b.race_ethnicity else a.race_ethnicity end) as race_ethnicity,a.headcount15,
b.headcount20 ,
( case when headcount20 is null then 0 else headcount20 end) -(case when headcount15 is null then 0 else headcount15 end) as headdiff
from hr20cwagg15 a
full outer join hr20cwagg20 b
on a.leadcoll=b.leadcoll and a.dept=b.dept and a.empgrp=b.empgrp and a.race_ethnicity=b.race_ethnicity;
quit;


/*generate macro for coll values from all three sources*/
proc sql;
create table tocfor as 
select substr(leadcoll,1,2) as leadcoll, 1 as sec
 from PPS
union
 select MAU, 2 as sec
 from AAprog

 union 
 select substr(Coll_1st,1,2) as leadColl,3 as sec
 from PAGset;
 quit;

 /**preparing for table of contents in proc document macro*/
 proc sort data=tocfor; by  leadcoll sec;
 data tocfor;
  set tocfor;
  seccount + 1;
  by leadcoll;
  if first.leadcoll then seccount = 1;
run;
data tocfor;
set tocfor;
counter+1;
run;

proc sql;
create table toc as 
select leadcoll, min(counter) as min, count(*) as cnt, min(sec) as minsec
from tocfor
group by leadcoll;

create table tocfor as
select  a.leadcoll,a.counter,min, cnt, minsec
from tocfor a
inner join toc b 
on a.leadcoll=b.leadcoll
where (a.seccount >1) or cnt=1;
quit;
/*excluding 43*/
data _null_;
set tocfor;
call symput('dpg'||left(_n_),counter);
call symput('rpg'||left(_n_),min);
where leadcoll ne 'Un';
run;

proc sql ;
select count(*) into :seccnt
from tocfor
where leadcoll ne 'Un';
quit;

/*TOC adjust pages macro*/
%macro pageadj;
%do i=1 %to &seccnt;
move Report#&&dpg&i.\Report#1 to report#&&rpg&i.;
 delete Report#&&dpg&i.;
%end;
%mend;

/*macro for loop over all colleges*/
data toc ; set toc; where leadcoll <> 'Un';

proc sql ;
select distinct count(distinct leadcoll)
 into :collcnt
 from toc;
quit;

data _null_;
set toc;
call symput('coll'||left(_n_),leadcoll);
call symput('minsec'||left(_n_),minsec);
run;

%put &minsec1;

data aaprog;
set aaprog;
dummy='AA';
college = substr(MAUName,1, find(MAUName,'-')-1);
run; 

data tst;
set pagset;
coll = upcase(substr(coll_1st,4, length(coll_1st)-4));
run;


proc format;
value empgrpfmt 1="Tenure System Faculty"
2='Fixed Term Faculty'
3='Other Academics';
run;

proc sql;
create table hrsupport as 
select distinct empgrp
from hr20cw;
quit;

data hrsupport;
set hrsupport;
length empgrplabel $35.;
if empgrp =4 then empgrplabel  = "Support Staff Headcount";
else if empgrp=1 then empgrplabel="Tenure System Faculty Headcount";
else if empgrp=2 then empgrplabel="Fixed Term Faculty Headcount";
else empgrplabel ='Other Academics Headcount';
run;
proc sql noprint;
select empgrp into: empgrp1- :empgrp4
from hrsupport;

select  empgrplabel into:empgrplabel1-:empgrplabel4
from hrsupport;
quit;
%put &empgrplabel1;
%put &version;

%macro ppsmetricdp;
	%do i=1 %to 1 /*&collcnt*/;
		data ds;
		set pps;
		where substr(leadcoll,1,2)="&&coll&i";
		call symput('collname',leadcoll);
		run;
proc sort data = ds; by dept; run;

data pagds;
		set pagset;
		coll = upcase(substr(coll_1st,4, length(coll_1st)-4));
		where substr(Coll_1st,1,2)="&&coll&i";
		call symput('collname',coll);
		run;
proc sort data = pagds; by Dept_1st; run;

data aads;
set aaprog;
where mau ="&&coll&i";
call symput('collname',COLLEGE);
run;

data hrds;
set hr20cw;
dummy1='D';
where substr(leadcoll,1,2)="&&coll&i";
		call symput('collname',leadcoll);
		run;

data hrdsw;
set hrw;
dummy='Y';
dummy1='D';
where substr(leadcoll,1,2)="&&coll&i";
		call symput('collname',leadcoll);
		run;


/*
%if &&minsec&i=1 %then %do;
ods proclabel="&collname.";
%end;
%else %do;
ods proclabel=' '; 
%end;*/

	
%do k=1 %to &varcnt;
%if &k <10 %then %do;
ods proclabel="PPS Metrics &collname."  ;
%end;
%else %do;
ods proclabel="Faculty and Academics- &collname."  ;
%end;
proc report data=ds style(report)={outputwidth=95% font_size=8pt} style(header)={font_face='Arial, Helvetica' font_size=8pt} contents="&&varlabel&k" ;
%if &k <10 %then %do;
title1 h=13pt "PPS Metrics - &&varlabel&k"  ;
%end;
%else %do;
title1 h=13pt "Faculty and Academics - &&varlabel&k"  ;
%end;
title2 h=6pt " ";
title3 h=10pt &collname ;
footnote3 h=9pt justify=left "PPS version: &version";
footnote4 j = r 'Page !{thispage} of !{lastpage}';
column dummy dept COL1,year_name ('Percent Change' Fiveyr Oneyr);
define dummy/group noprint;
define dept/group 'Department' ;
define year_name/'Year' across;
%if &k= 5 %then %do;
define COL1/' ' analysis sum format=dollar15.0 ;
%end;
%else %if &k= 7 %then %do;
define COL1/' ' analysis sum format=percentn15.0 ;
%end;
%else %do;
define COL1/' ' analysis sum format=comma15.0 ;
%end;
define fiveyr/'Five Year' computed format=percentn15.1;
define Oneyr/'One Year' computed format=percentn15.1;
compute fiveyr;
fiveyr = (_c8_- _c3_)/_c3_;
%if &k= 5 %then %do;
if fiveyr lt &&ppspctfiveyr&k & fiveyr ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if fiveyr gt &&ppspctfiveyr&k & fiveyr ne . then call define(_col_,"style","style={background=#fff7bc}");
	
%end;
%else %do;
if fiveyr gt &&ppspctfiveyr&k & fiveyr ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if fiveyr lt &&ppspctfiveyr&k & fiveyr ne . then call define(_col_,"style","style={background=#fff7bc}");
	 
%end;
endcomp;
compute oneyr;
oneyr = (_c8_- _c7_)/_c7_;
%if &k= 5 %then %do;
if oneyr lt &&ppspctoneyr&k & oneyr ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if oneyr gt &&ppspctoneyr&k & oneyr ne . then call define(_col_,"style","style={background=#fff7bc}");
	 

%end;
%else %do;
if oneyr gt &&ppspctoneyr&k & oneyr ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if oneyr lt &&ppspctoneyr&k & oneyr ne . then call define(_col_,"style","style={background=#fff7bc}");
	
%end;
endcomp;
compute dept;
     if find(Dept,'All','i') then call define(_row_,"style","style={font_weight=bold}");
	
 endcomp;
 
break before dummy/ page contents=""; /*remove 3rd node*/
where var="&&var&k";
run;
%end; 

/**faculty race*/
%do j=1 %to 4;

ods proclabel="&&empgrplabel&j Race/Ethnicity Breakdown 2020"  ;
proc report data= hrds style(report)={outputwidth=95% font_size=8pt} style(header)={font_face='Arial, Helvetica' font_size=8pt} contents="&collname" ;
title1 h=13pt "&&empgrplabel&j ,Year of 2020"  ;
title2 h=6pt " ";
title3 h=10pt &collname ;
column dummy1 dummy dept  headcount,race_ethnicity ('Total' headcount=tot);
define dummy1/group noprint;
define dummy/group noprint;
define dept/ order group 'Department';
define race_ethnicity/' ' across;
define tot / analysis sum ' ' f=comma6. ;
break before dummy/ summarize ;
compute before dummy;
 dept="&collname -All";
 call define(_row_,"style","style={font_weight=bold}");
endcomp;
break before dummy1/ page contents=""  ;
where empgrp =&&empgrp&j and cyear='2020';
run;


ods proclabel="&&empgrplabel&j Race/Ethnicity Breakdown 2015"  ;
proc report data= hrds style(report)={outputwidth=95% font_size=8pt} style(header)={font_face='Arial, Helvetica' font_size=8pt} contents="&collname" ;
title1 h=13pt "&&empgrplabel&j ,Year of 2015"  ;
title2 h=6pt " ";
title3 h=10pt &collname ;
column dummy1 dummy dept  headcount,race_ethnicity ('Total' headcount=tot);
define dummy1/group noprint;
define dummy/group noprint;
define dept/ order group 'Department';
define race_ethnicity/' ' across;
define tot / analysis sum ' ' f=comma6. ;
break before dummy/ summarize ;
compute before dummy;
 dept="&collname -All";
 call define(_row_,"style","style={font_weight=bold}");
endcomp;
break before dummy1/ page contents=""  ;
where empgrp =&&empgrp&j and cyear='2015';
run;


ods proclabel="&&empgrplabel&j Race/Ethnicity Breakdown Five Year Difference"  ;
proc report data= hrdsw style(report)={outputwidth=95% font_size=8pt} style(header)={font_face='Arial, Helvetica' font_size=8pt} contents="&collname" ;
title1 h=13pt "&&empgrplabel&j Five Year Difference "  ;
title2 h=6pt " ";
title3 h=10pt &collname ;
column dummy1 dummy dept  race_ethnicity, ( headdiff  test) ('Total' tot);
define dummy1/group noprint;
define dummy/group noprint;
define dept/ order group 'Department';
define race_ethnicity/' ' across;
define headdiff/ Analysis Sum noprint;
define test/'' computed;
define tot / computed ' ' f=comma6. ;
compute test;
_c5_= _c4_;
_c7_= _c6_;
_c9_= _c8_;
_c11_= _c10_;
_c13_= _c12_;
_c15_= _c14_;

if _c4_ gt 0 & _c4_ ne . then call define("_c5_","style","style={background=#a1d99b}");
	 else if _c4_ lt 0 & _c4_ ne . then call define("_c5_","style","style={background=#fff7bc}");	

if _c6_ gt 0 & _c6_ ne . then call define("_c7_","style","style={background=#a1d99b}");
	 else if _c6_ lt 0 & _c6_ ne . then call define("_c7_","style","style={background=#fff7bc}");	

if _c8_ gt 0 & _c8_ ne . then call define("_c9_","style","style={background=#a1d99b}");
	 else if _c8_ lt 0 & _c8_ ne . then call define("_c9_","style","style={background=#fff7bc}");	

if _c10_ gt 0 & _c10_ ne . then call define("_c11_","style","style={background=#a1d99b}");
	 else if _c10_ lt 0 & _c10_ ne . then call define("_c11_","style","style={background=#fff7bc}");	

if _c12_ gt 0 & _c12_ ne . then call define("_c13_","style","style={background=#a1d99b}");
	 else if _c12_ lt 0 & _c12_ ne . then call define("_c13_","style","style={background=#fff7bc}");

if _c14_ gt 0 & _c14_ ne . then call define("_c15_","style","style={background=#a1d99b}");
	 else if _c14_ lt 0 & _c14_ ne . then call define("_c15_","style","style={background=#fff7bc}");	
	
endcomp;	
compute tot;
tot= sum(_c5_,_c7_,_c9_,_c11_,_c13_,_c15_);
if tot gt 0 & tot ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if tot lt 0 & tot ne . then call define(_col_,"style","style={background=#fff7bc}");	
 
endcomp;

break before dummy/ summarize ;

compute before dummy;
 dept="&collname -All";
 call define(_row_,"style","style={font_weight=bold}");
endcomp;
break before dummy1/ page contents=""  ;
where empgrp =&&empgrp&j ;
run;
%end;



/*PAG*/


%do k=1 %to &pagvarcnt;

ods proclabel="PAG Metrics &collname."  ;
proc report data=pagds  style(report)={outputwidth=95% font_size=8pt} style(header)={font_face='Arial, Helvetica' font_size=8pt} contents="&&paglabel&k" ;
title1 h=13pt "PAG Metrics - &&paglabel&k"  ;
title2 h=6pt " ";
title3 h=10pt &collname ;
footnote3 h=9pt justify=left "PAG version: SISFull(FS20)";
footnote4 j = r 'Page !{thispage} of !{lastpage}';
column dummy dept_1st COL1,cohort1 ('Percent Change' Fiveyr Oneyr);
define dummy/group noprint;
%if &k= 7 %then %do;
define dept_1st/group 'Degree Department' ;
%end;
%else %do;
define dept_1st/group 'Entering Department' ;
%end;
%if &k= 7 %then %do;
define cohort1/'Graduating Cohort' across;
%end;
%else %do;
define cohort1/'Entering Cohort' across;
%end;
%if &k= 7 %then %do;
define COL1/' ' analysis sum format=comma15.2;
%end;
%else %do;
define COL1/' ' analysis sum format=percentn15.1 ;
%end;
define fiveyr/'Five Year' computed format=percentn15.1;
define Oneyr/'One Year' computed format=percentn15.1;

compute fiveyr;
fiveyr = (_c8_- _c3_)/_c3_;
%if &k=7 %then %do;
if fiveyr lt &&pagpctfiveyr&k & fiveyr ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if fiveyr gt &&pagpctfiveyr&k & fiveyr ne . then call define(_col_,"style","style={background=#fff7bc}");
	
endcomp;
%end;
%else %do;
if fiveyr gt &&pagpctfiveyr&k & fiveyr ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if fiveyr lt &&pagpctfiveyr&k & fiveyr ne . then call define(_col_,"style","style={background=#fff7bc}");
	
endcomp;
%end;
compute oneyr;
oneyr = (_c8_- _c7_)/_c7_;
%if &k=7 %then %do;
if oneyr lt &&pagpctoneyr&k & oneyr ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if oneyr gt &&pagpctoneyr&k & oneyr ne . then call define(_col_,"style","style={background=#fff7bc}");
	
%end;
%else %do;
if oneyr gt &&pagpctoneyr&k & oneyr ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if oneyr lt &&pagpctoneyr&k & oneyr ne . then call define(_col_,"style","style={background=#fff7bc}");
	 
%end;
endcomp;

 compute dept_1st;
     if  find(Dept_1st,'All','i') then call define(_row_,"style","style={font_weight=bold}");
	
 endcomp;
break before dummy/ page contents=""; /*remove 3rd node*/

where _NAME_="&&pagvar&k" ;
run;
%end;

/*AA*/
ods proclabel="AA Metrics &collname."  ;
proc report data=aads style(report)={outputwidth=95% font_size=8pt} style(header)={font_face='Arial, Helvetica' font_size=8pt}
contents="SRI" spanrows ;
title1 h=13pt "AA Metrics -SRI"  ;
title2 h=6pt " ";
title3 h=10pt &collname ;
footnote3 h=9pt justify=left " ";
footnote4 j = r 'Page !{thispage} of !{lastpage}';
column dummy Dept TAXONOMY  PCT_SCHLR_RSCH_INDEX1,year ('Percent Change'   fiveyr Oneyr);
define dummy/ group noprint;
define Dept/ order group 'Department' ;
define TAXONOMY/group 'Taxonomy';
define year/'Year' across;
define PCT_SCHLR_RSCH_INDEX1/' ' analysis sum format=comma15.2 ;
define fiveyr/'Five Year' computed format=percentn15.1;
define Oneyr/'One Year' computed format=percentn15.1;
compute fiveyr;
fiveyr = (_c6_- _c4_)/_c4_;
if fiveyr gt 0 & fiveyr ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if fiveyr lt 0 & fiveyr ne . then call define(_col_,"style","style={background=#fff7bc}");	
endcomp;
compute oneyr;
oneyr = (_c6_- _c5_)/_c5_;
if oneyr gt 0 & oneyr ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if oneyr lt 0 & oneyr ne . then call define(_col_,"style","style={background=#fff7bc}");	
endcomp;
break before dummy/ page contents=""; /*remove 3rd node*/
run;


%end;
%mend ppsmetricdp;





ods listing close;
/**output prep for proc document*/
ods document name=mydocrep(write);
%ppsmetricdp;
ods document close;


proc document name=mydocrep ;

run;
options   nodate  nonumber leftmargin=0.5in rightmargin=0.5in orientation=landscape papersize=A4 center missing=' ' ;

ods pdf file="O:\IS\Internal\Reporting\Annual Reports\Key Metrics Set\Output\TTKeyMetrics_CollegeDept_Trend&todaysDate..pdf" style=Custom contents=yes;
ods escapechar= '!';
 replay;
run;
ods listing;
ods pdf close;
quit; 
