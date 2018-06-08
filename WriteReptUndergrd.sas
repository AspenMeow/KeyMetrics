/***Connect to db***/
%include "H:\SAS\SAS_Log_EDW.sas" ;
LIBNAME BIMSUTST  odbc required="DSN=BIMSUTST; uid=&BIMSUTST_uid; pwd=&BIMSUTST_pwd";
libname PAG oracle path="MSUEDW" user="&MSUEDW_uid" pw= "&MSUEDW_pwd"
schema = OPB_PERS_FALL preserve_tab_names = yes connection=sharedread;

/*********PPS metrics import pre processing from R ppskeymetricsds.R*************/
/********PPS metrics always using the latest year available*********************/

proc import datafile="O:\IS\Internal\Reporting\Annual Reports\Key Metrics Set\Data\PPSmetrics.xlsx"  out=PPS dbms=xlsx replace ;sheet='PPS';
run;
/**PPS quantitle*
proc import datafile="O:\IS\Internal\Reporting\Annual Reports\Key Metrics Set\Data\PPSmetrics.xlsx"  out=PPSqt dbms=xlsx replace ;sheet='PPSquantitle';
run;*/

/*college with grad prof students*/
proc sql ;
create table pps  as 
select a.*, coll.gradprofcnt
from pps a 
left join
      ( 
	select MANAGE_LEVEL_3, max(Grad_Prof) as gradprofcnt
		from pps
		group by MANAGE_LEVEL_3
		) coll
on a.MANAGE_LEVEL_3=coll.MANAGE_LEVEL_3
where a.Undergraduates ne . and a.Undergraduates>0
order by a.MANAGE_LEVEL_3,a.orderdp,a.LEVEL4;

quit;




/*generate macro variable for PPS qt**/
/*proc sql ;
 select distinct value
 into :varqt1-:varqt10 
 from PPSqt
 order by vp;
 quit;*/

data dat ;
set pps;
where orderdp=0;
run;

PROC UNIVARIATE data =dat;
where orderdp=0;
 var V1 V2 V3 V4 V5;
 output  out=qt pctlpre= PSV1 PSV2 PSV3 PSV4 PSV5 pctlpts=33.3 66.6;
 run;

 proc transpose data= qt out=qtlg;
run;
data _null_;
set qtlg;
call symput(_NAME_,COL1);
run;


/*get PPS yrs as macro*/
 data _null_;
 set PPS;
 call symput('Ungrdyr',Undergraduates_yr);
 call symput('Masteryr',Masters_yr);
 call symput('Docyr',Doctoral_yr);
 call symput('Profyr',Grad_Prof_yr);
 call symput('V1yr',V1_yr);
  call symput('V2yr',V2_yr);
   call symput('V3yr',V3_yr);
  call symput('V4yr',V4_yr);
 call symput('V5yr',V5_yr);
run;
%put &V5yr;
%put %substr(&Ungrdyr,1,4);




 /*get PPS version macro var*/
proc sql noprint;
select VERDATE
into :version
from BIMSUTST.version;
quit;


/*********************************************************************/
/******AA data******************/
%let AAver='06.887';
proc sql stimer;
/*AAprogram level data*/
create table AAprog as
select *
from BIMSUTST.ACADEMIC_ANALYTICS_METRICS_P
where VERSION=&AAver;

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
order by MAU, PHD_Program;
quit;

/*get 1/3 and 2/3 percentitle cutoff value*/
PROC UNIVARIATE data =AAprog;
 var PCT_JOURNAL_PUBS_PER_FACULTY1 PCT_CITATIONS_PER_FACULTY1 PCT_CITATIONS_PER_PUBLICATION1 PCT_GRANTS_PER_FACULTY1
PCT_GRANT_DOLLARS_PER_FACULTY1 PCT_DOLLARS_PER_GRANT1 PCT_AWARDS_PER_FACULTY1;
 output  out=ds pctlpre=JP CF CP GF GDF DG AF pctlpts=33.3 66.6;
 run;
proc transpose data= ds out=lg;
run;
/*create macro variables for percentitle cutoff aa*/
data _null_;
set lg;
call symput(_NAME_,COL1);
run;



/************************************************/
/****************PAG*******************************/
%let entrycohort =2011;
%let gradcohort ='2016-2017';


proc sql stimer;
create table PAG_Enter_Dp as 
select COHORT, strip(COLLEGE_FIRST)||'-'|| strip(COLLEGE_FIRST_NAME) as Coll_1st, strip(DEPT_FIRST) ||'-'|| strip(DEPT_FIRST_NAME) as Dept_1st,
       count(PID) as N, avg(PERSIST1) as PERSIST1,avg(PERSIST2) as PERSIST2,
	   avg(PERSIST3) as PERSIST3, avg(GRAD4) as GRAD4,
	   avg(GRAD5)as GRAD5, avg(GRAD6) as GRAD6
from PAG.PERSISTENCE_V
where Student_LEVEL='UN'
and (  ENTRANT_SUMMER_FALL='Y' or substr(ENTRY_TERM_CODE,1,1)='F' )
and LEVEL_ENTRY_STATUS='FRST'
and cohort = &entrycohort
group by COHORT, COLLEGE_FIRST, COLLEGE_FIRST_NAME, DEPT_FIRST, DEPT_FIRST_NAME;

create table PAG_Enter_Coll as 
select COHORT, strip(COLLEGE_FIRST)||'-'|| strip(COLLEGE_FIRST_NAME) as Coll_1st, strip(COLLEGE_FIRST)||'-'|| strip(COLLEGE_FIRST_NAME)||'-All' as Dept_1st,
       count(PID) as N, avg(PERSIST1) as PERSIST1,avg(PERSIST2) as PERSIST2,
	   avg(PERSIST3) as PERSIST3, avg(GRAD4) as GRAD4,
	   avg(GRAD5)as GRAD5, avg(GRAD6) as GRAD6
from PAG.PERSISTENCE_V
where Student_LEVEL='UN'
and (  ENTRANT_SUMMER_FALL='Y' or substr(ENTRY_TERM_CODE,1,1)='F' )
and LEVEL_ENTRY_STATUS='FRST'
and cohort = &entrycohort
group by COHORT, COLLEGE_FIRST, COLLEGE_FIRST_NAME;




create table PAG_Grad_Dp as 
select GRADUATING_COHORT, strip(COLLEGE_DEGREE)||'-'|| strip(COLLEGE_DEGREE_NAME) as Coll_DEGR, strip(DEPT_DEGREE) ||'-'|| strip(DEPT_DEGREE_NAME) as Dept_DEGR,
       count(PID) as N,AVG(TTD_IN_YEARS) AS TTD
from PAG.PERSISTENCE_V
where Student_LEVEL='UN'
and (  ENTRANT_SUMMER_FALL='Y' or substr(ENTRY_TERM_CODE,1,1)='F' )
and LEVEL_ENTRY_STATUS='FRST'
and Graduating_Cohort= &gradcohort
group by GRADUATING_COHORT, COLLEGE_DEGREE, COLLEGE_DEGREE_NAME, DEPT_DEGREE, DEPT_DEGREE_NAME;

create table PAG_Grad_Coll as 
select GRADUATING_COHORT, strip(COLLEGE_DEGREE)||'-'|| strip(COLLEGE_DEGREE_NAME) as Coll_DEGR, strip(COLLEGE_DEGREE)||'-'|| strip(COLLEGE_DEGREE_NAME)||'-All' as Dept_DEGR,
       count(PID) as N,AVG(TTD_IN_YEARS) AS TTD
from PAG.PERSISTENCE_V
where Student_LEVEL='UN'
and (  ENTRANT_SUMMER_FALL='Y' or substr(ENTRY_TERM_CODE,1,1)='F' )
and LEVEL_ENTRY_STATUS='FRST'
and Graduating_Cohort= &gradcohort
group by GRADUATING_COHORT, COLLEGE_DEGREE, COLLEGE_DEGREE_NAME;

run;

/*combine ds*/
data PAG_enter;
set PAG_enter_Dp PAG_enter_Coll;
run;
 
data PAG_Degr;
set PAG_Grad_Dp PAG_Grad_Coll;
run;

/**compute PAG quantitle *****/
PROC UNIVARIATE data =PAG_enter_Dp;
 var PERSIST1 PERSIST2 PERSIST3 GRAD4 GRAD5 GRAD6;
 output  out=pagds pctlpre=P1 P2 P3 G4 G5 G6 pctlpts=33.3 66.6;
 run;
proc transpose data= pagds out=paglg;
run;
/*create macro variables for percentitle cutoff pag entering cohort*/
data _null_;
set paglg;
call symput(_NAME_,COL1);
run;

/*create macro variable for TTD*/
PROC UNIVARIATE data =PAG_Grad_Dp;
var TTD;
 output  out=ttds pctlpre=TTD pctlpts=33.3 66.6;
 run;
 
data _null_;
set ttds;
call symput('TTD33_3', TTD33_3);
call symput('TTD66_6',TTD66_6);
run;


/****************************************************/
/*generate macro for coll values from all three sources*/
proc sql stimer;
 create table allcoll as 
 select MANAGE_LEVEL_3
 from PPS
 union
 select MAU
 from AAprog
 union 
 select substr(Coll_1st,1,2) as Coll
 from PAG_enter
 union 
 select substr(Coll_degr,1,2) as Coll
 from PAG_Degr;

 select distinct count(distinct MANAGE_LEVEL_3)
 into :collcnt
 from allcoll;

select distinct MANAGE_LEVEL_3
 into :coll1-:coll%left(&collcnt)
 from allcoll
 order by MANAGE_LEVEL_3;
 quit;






/*******set cutomized report template******************/


proc template;
 define style Styles.Custom;
 parent = Styles.Printer;
 Style SystemTitle from SystemTitle /
Font = ("Arial, Helvetica", 6pt, bold );
Style SystemFooter from systemFooter /
Font = (“Arial, Helvetica”, 6pt, bold );
 
replace color_list /
'link' = blue /* links */
/*'bgH' = grayBB  row and column header background */
'bgT' = white /* table background */
'bgD' = white /* data cell background */
'fg' = black /* text color */
'bg' = white; 
replace Table from Output /
 frame = box /* outside borders: void, box, above/below, vsides/hsides, lhs/rhs */
rules = all /* internal borders: none, all, cols, rows, groups */
cellpadding = 4pt /* the space between table cell contents and the cell border */
cellspacing = 0.2pt /* the space between table cells, allows background to show */
borderwidth = 0.1pt /* the width of the borders and rules */
bordercolor = #636363
background = color_list('bgT') /* table background color */;
 end;
run;










%macro ppsmetric;
	%do k=1 %to &collcnt;
		data ds;
		set pps;
		where MANAGE_LEVEL_3="&&coll&k";
		call symput('collname',leadcoll);
		call symput('gradprofn',gradprofcnt);
		run;

		
		%if &gradprofn=0 %then %do;
		proc report data= ds  style(report)={outputwidth=100%} style(header)={font_face='Arial, Helvetica' font_size=8pt}  ;
		title1 h=13pt "PPS Metrics by Lead Organization Structure"  ;
		title2 h=6pt "";
		title3 "&collname" ;
		footnote1 h=9pt justify=left "Note: master student counts include non-degree seeking lifelong graduates and MSU Law lifelong students." ;
		footnote2 h=9pt justify=left "Departments included are those with undergraduates and led by instructional colleges and have data in at least one PPS metrics.";
		footnote3 h=9pt justify=left "PPS version: &version";
		column Dept Undergraduates Masters Doctoral V1 V2 V3 V4 V5;
		define Dept/display 'Dept Name'   ;
		define Undergraduates/display "Undergrads Fall %substr(&Ungrdyr,1,4)" format=comma15.0 ;
		define Masters/display "Masters Fall %substr(&Masteryr,1,4)" format=comma15.0 ;
		define Doctoral/display "Doctoral Fall %substr(&Masteryr,1,4)" format=comma15.0 ;
		define V1/display "General Fund Budget/FY Admin SCH &V1yr" format=dollar15.0 style(column)=[width=8%];
		define V2/display "Admin Based SCH per Rank Fac FY &V2yr" format=comma15.0 style(column)=[width=8%];
		define V3/display "Tuition Revenue/Instructional Cost &V3yr" format=percentn15.0 style(column)=[width=8%];
		define V4/display "Total Grants - 3 Year Average-PI-Real$ &V4yr" format=dollar15.0;
		define V5/display "Total Grant PI Based/TS AF-FTE &V5yr" format=dollar15.0 style(column)=[width=8%];

 compute V1;
     if V1 lt &PSV133_3 & V1 ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if V1 gt &PSV166_6 & V1 ne . then call define(_col_,"style","style={background=#fff7bc}");
	 else if V1 ne . then call define(_col_,"style","style={background=#bdbdbd}");
	 if find(_c1_,'All','i') then  CALL DEFINE("_c5_", "style","STYLE=[BACKGROUND=white]");
	
 endcomp;
 compute V2;
     if V2 lt &PSV233_3 & V2 ne . then call define(_col_,"style","style={background=#fff7bc}");
	 else if V2 gt &PSV266_6 & V2 ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if V2 ne . then call define(_col_,"style","style={background=#bdbdbd}");
	 if find(_c1_,'All','i') then  CALL DEFINE("_c6_", "style","STYLE=[BACKGROUND=#ffffff]");
	
 endcomp;
 compute V3;
     if V3 lt &PSV333_3 & V3 ne . then call define(_col_,"style","style={background=#fff7bc}");
	 else if V3 gt &PSV366_6 & V3 ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if V3 ne . then call define(_col_,"style","style={background=#bdbdbd}");
	 if find(_c1_,'All','i') then  CALL DEFINE("_c7_", "style","STYLE=[BACKGROUND=#ffffff]");
	
 endcomp;
 compute V4;
     if V4 lt &PSV433_3 & V4 ne . then call define(_col_,"style","style={background=#fff7bc}");
	 else if V4 gt &PSV466_6 & V4 ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if V4 ne . then call define(_col_,"style","style={background=#bdbdbd}");
	 if find(_c1_,'All','i') then  CALL DEFINE("_c8_", "style","STYLE=[BACKGROUND=#ffffff]");
	
 endcomp;
 compute V5;
     if V5 lt &PSV533_3 & V5 ne . then call define(_col_,"style","style={background=#fff7bc}");
	 else if V5 gt &PSV566_6 & V4 ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if V5 ne . then call define(_col_,"style","style={background=#bdbdbd}");
	 if find(_c1_,'All','i') then  CALL DEFINE("_c9_", "style","STYLE=[BACKGROUND=#ffffff]");
	
 endcomp;
 compute Dept;
     if find(Dept,'All','i') then call define(_row_,"style","style={font_weight=bold}");
	
 endcomp;
 
run;
%end;
%else %do;
proc report data= ds  style(report)={outputwidth=100%} style(header)={font_face='Arial, Helvetica' font_size=8pt};
		title1 h=13pt "PPS Metrics by Lead Organization Structure"  ;
		title2 h=6pt "";
		title3 "&collname" ;
		footnote1 h=9pt justify=left "Note: master student counts include non-degree seeking lifelong graduates and MSU Law lifelong students." ;
		footnote2 h=9pt justify=left "Departments included are those led by instructional colleges and have data in at least one PPS metrics.";
		footnote3 h=9pt justify=left "PPS version: &version";
		column Dept Undergraduates Masters Doctoral Grad_Prof V1 V2 V3 V4 V5;
		define Dept/display 'Dept Name' ;
		define Undergraduates/display "Undergrads Fall %substr(&Ungrdyr,1,4)" format=comma15.0 ;
		define Masters/display "Masters Fall %substr(&Masteryr,1,4)" format=comma15.0 ;
		define Doctoral/display "Doctoral Fall %substr(&Masteryr,1,4)" format=comma15.0 ;
		define Grad_Prof/display "Professional Students Fall %substr(&Masteryr,1,4)" format=comma15.0 ;
		define V1/display "General Fund Budget/FY Admin SCH &V1yr" format=dollar15.0 style(column)=[width=8%];
		define V2/display "Admin Based SCH per Rank Fac FY &V2yr" format=comma15.0 style(column)=[width=8%];
		define V3/display "Tuition Revenue/Instructional Cost &V3yr" format=percentn15.0 style(column)=[width=8%];
		define V4/display "Total Grants - 3 Year Average-PI-Real$ &V4yr" format=dollar15.0;
		define V5/display "Total Grant PI Based/TS AF-FTE &V5yr" format=dollar15.0 style(column)=[width=8%];

 compute V1;
     if V1 lt &PSV133_3 & V1 ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if V1 gt &PSV166_6 & V1 ne . then call define(_col_,"style","style={background=#fff7bc}");
	 else if V1 ne . then call define(_col_,"style","style={background=#bdbdbd}");
	 if find(_c1_,'All','i') then  CALL DEFINE("_c6_", "style","STYLE=[BACKGROUND=white]");
	
 endcomp;
 compute V2;
     if V2 lt &PSV233_3 & V2 ne . then call define(_col_,"style","style={background=#fff7bc}");
	 else if V2 gt &PSV266_6 & V2 ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if V2 ne . then call define(_col_,"style","style={background=#bdbdbd}");
	 if find(_c1_,'All','i') then  CALL DEFINE("_c7_", "style","STYLE=[BACKGROUND=#ffffff]");
	
 endcomp;
 compute V3;
     if V3 lt &PSV333_3 & V3 ne . then call define(_col_,"style","style={background=#fff7bc}");
	 else if V3 gt &PSV366_6 & V3 ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if V3 ne . then call define(_col_,"style","style={background=#bdbdbd}");
	 if find(_c1_,'All','i') then  CALL DEFINE("_c8_", "style","STYLE=[BACKGROUND=#ffffff]");
	
 endcomp;
 compute V4;
     if V4 lt &PSV433_3 & V4 ne . then call define(_col_,"style","style={background=#fff7bc}");
	 else if V4 gt &PSV466_6 & V4 ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if V4 ne . then call define(_col_,"style","style={background=#bdbdbd}");
	 if find(_c1_,'All','i') then  CALL DEFINE("_c9_", "style","STYLE=[BACKGROUND=#ffffff]");
	
 endcomp;
 compute V5;
     if V5 lt &PSV533_3 & V5 ne . then call define(_col_,"style","style={background=#fff7bc}");
	 else if V5 gt &PSV566_6 & V4 ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if V5 ne . then call define(_col_,"style","style={background=#bdbdbd}");
	 if find(_c1_,'All','i') then  CALL DEFINE("_c10_", "style","STYLE=[BACKGROUND=#ffffff]");
	
 endcomp;
 compute Dept;
     if find(Dept,'All','i') then call define(_row_,"style","style={font_weight=bold}");
	
 endcomp;
 
run;
%end;	

/**AA**/
data aa;
set AAprog;
MAUName= MAU||'-'|| MAUName;
where MAU="&&coll&k";
call symput('MAUName',MAUName);
call symput('AAver', VERSION);
call symput('Yr',YEAR );
run;

proc report data= aa  split='~'  spanrows style(report)={outputwidth=100%}  style(header)={font_face='Arial, Helvetica' font_size=8pt};
title1 h=13pt "Academic Analytics Metrics PhD Program View &Yr" ;
title2 h=6pt "";
title3 &MAUName ;
footnote1 h=9pt justify=left "Data Source: Academic Analytics Program View" ;
footnote2 h=9pt justify=left "AA version: &AAver";
footnote3 justify=left " ";
column PHD_PROGRAM FACULTY_IN_PROG NATIONAL_ACADEMY_MEMBERS PROG_IN_MULT_TAXONOMIES TAXONOMY 
	AAU_PROG_IN_DISCIPLINE1 PCT_SCHLR_RSCH_INDEX1  RANK_SCHLR_RSCH_INDEX1 PCT_JOURNAL_PUBS_PER_FACULTY1 PCT_CITATIONS_PER_FACULTY1
 PCT_CITATIONS_PER_PUBLICATION1 PCT_GRANTS_PER_FACULTY1 PCT_GRANT_DOLLARS_PER_FACULTY1 PCT_DOLLARS_PER_GRANT1 PCT_AWARDS_PER_FACULTY1;
define PHD_PROGRAM/ order 'PhD Program'  style(column)=[vjust=m just=l ] ;
define FACULTY_IN_PROG/ order 'N of Faculty' style(column)=[vjust=m   ];
define NATIONAL_ACADEMY_MEMBERS/ order 'N of National Academy Member' style(column)=[vjust=m   ];
define PROG_IN_MULT_TAXONOMIES/ order 'Program Mutiple Taxonomy' style(column)=[vjust=m  ];
define TAXONOMY/display   'Taxonomy' style(column)=[cellwidth=10%]  ;
define AAU_PROG_IN_DISCIPLINE1/display 'N of Program in discipline at AAU' ;
define PCT_SCHLR_RSCH_INDEX1/display 'FSPI Percentitle' format=comma10.2 ;
define RANK_SCHLR_RSCH_INDEX1/display 'FPSI Rank' format=comma10.0 ;
define PCT_JOURNAL_PUBS_PER_FACULTY1/display 'Percentile on Journal Pubs per faculty' format=comma10.2  ;
define PCT_CITATIONS_PER_FACULTY1/display 'Percentile on Citations per faculty' format=comma10.2 ;
define PCT_CITATIONS_PER_PUBLICATION1/display 'Percentile on Citations per publication' format=comma10.2 ;
define PCT_GRANTS_PER_FACULTY1/display 'Percentile on Grants per faculty' format=comma10.2  ;
define PCT_GRANT_DOLLARS_PER_FACULTY1/display 'Percentitle on Grant Dollars per faculty' format=comma10.2 ;
define PCT_DOLLARS_PER_GRANT1/display 'Percentitle Dollars per Grant' format=comma10.2 ;
define PCT_AWARDS_PER_FACULTY1/display 'Percentitle on Awards per faculty' format=comma10.2 ;
compute PCT_JOURNAL_PUBS_PER_FACULTY1;
     if PCT_JOURNAL_PUBS_PER_FACULTY1 lt &JP33_3 & PCT_JOURNAL_PUBS_PER_FACULTY1 ne . then call define(_col_,"style","style={background= #fff7bc}");
	 else if PCT_JOURNAL_PUBS_PER_FACULTY1 gt &JP66_6 & PCT_JOURNAL_PUBS_PER_FACULTY1 ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if PCT_JOURNAL_PUBS_PER_FACULTY1 ne . then call define(_col_,"style","style={background=#bdbdbd}");
 endcomp;
 compute PCT_CITATIONS_PER_FACULTY1;
     if PCT_CITATIONS_PER_FACULTY1 lt &CF33_3 & PCT_CITATIONS_PER_FACULTY1 ne . then call define(_col_,"style","style={background= #fff7bc}");
	 else if PCT_CITATIONS_PER_FACULTY1 gt &CF66_6 & PCT_CITATIONS_PER_FACULTY1 ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if PCT_CITATIONS_PER_FACULTY1 ne . then call define(_col_,"style","style={background=#bdbdbd}");
 endcomp;
 compute PCT_CITATIONS_PER_PUBLICATION1;
     if PCT_CITATIONS_PER_PUBLICATION1 lt &CP33_3 & PCT_CITATIONS_PER_PUBLICATION1 ne . then call define(_col_,"style","style={background= #fff7bc}");
	 else if PCT_CITATIONS_PER_PUBLICATION1 gt &CP66_6 & PCT_CITATIONS_PER_PUBLICATION1 ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if PCT_CITATIONS_PER_PUBLICATION1 ne . then call define(_col_,"style","style={background=#bdbdbd}");
 endcomp;
  compute PCT_GRANTS_PER_FACULTY1;
     if PCT_GRANTS_PER_FACULTY1 lt &GF33_3 & PCT_GRANTS_PER_FACULTY1 ne . then call define(_col_,"style","style={background= #fff7bc}");
	 else if PCT_GRANTS_PER_FACULTY1 gt &GF66_6 & PCT_GRANTS_PER_FACULTY1 ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if PCT_GRANTS_PER_FACULTY1 ne . then call define(_col_,"style","style={background=#bdbdbd}");
 endcomp;
  compute PCT_GRANT_DOLLARS_PER_FACULTY1;
     if PCT_GRANT_DOLLARS_PER_FACULTY1 lt &GDF33_3 & PCT_GRANT_DOLLARS_PER_FACULTY1 ne . then call define(_col_,"style","style={background= #fff7bc}");
	 else if PCT_GRANT_DOLLARS_PER_FACULTY1 gt &GDF66_6 & PCT_GRANT_DOLLARS_PER_FACULTY1 ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if PCT_GRANT_DOLLARS_PER_FACULTY1 ne . then call define(_col_,"style","style={background=#bdbdbd}");
 endcomp;
  compute PCT_DOLLARS_PER_GRANT1;
     if PCT_DOLLARS_PER_GRANT1 lt &DG33_3 & PCT_DOLLARS_PER_GRANT1 ne . then call define(_col_,"style","style={background= #fff7bc}");
	 else if PCT_DOLLARS_PER_GRANT1 gt &DG66_6 & PCT_DOLLARS_PER_GRANT1 ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if PCT_DOLLARS_PER_GRANT1 ne . then call define(_col_,"style","style={background=#bdbdbd}");
 endcomp;
  compute PCT_AWARDS_PER_FACULTY1;
     if PCT_AWARDS_PER_FACULTY1 lt &AF33_3 & PCT_AWARDS_PER_FACULTY1 ne . then call define(_col_,"style","style={background= #fff7bc}");
	 else if PCT_AWARDS_PER_FACULTY1 gt &AF66_6 & PCT_AWARDS_PER_FACULTY1 ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if PCT_AWARDS_PER_FACULTY1 ne . then call define(_col_,"style","style={background=#bdbdbd}");
 endcomp;
 
run;

/**PAGenter**/
data PAG;
set PAG_enter;
MAU= substr(Coll_1st,1,2);
if MAU="&&coll&k";
call symput('entercoll',Coll_1st);
run;

proc report data= pag  split='~'  spanrows style(report)={outputwidth=100%} style(header)={font_face='Arial, Helvetica' font_size=8pt} ;
title1 h=13pt "First-Time Undergraduates Persistence At and Graduation Rate From MSU By First Department" ;
title2 h=6pt "";
title3 h=12pt &entercoll ;
title4 "Entering Cohort : &entrycohort ";
footnote1 h=9pt justify=left "Data Source: PAG " ;
footnote2 h=9pt justify=left " ";
column Dept_1st N PERSIST1 PERSIST2 PERSIST3 GRAD4 GRAD5 GRAD6;
define Dept_1st/display 'Entering Department'  style(column)=[vjust=m just=l ] ;
define N/ display 'N'  format=comma10.0;
define PERSIST1/ display 'Persist 1st Returning Fall' format=comma10.1 ;
define PERSIST2/ display 'Persist 2nd Returning Fall' format=comma10.1 ;
define PERSIST3/ display 'Persist 3rd Returning Fall' format=comma10.1 ;
define Grad4/ display 'Graduated by 4th Yr' format=comma10.1 ;
define Grad5/ display 'Graduated by 5th Yr' format=comma10.1 ;
define Grad6/ display 'Graduated by 6th Yr' format=comma10.1 ;
compute PERSIST1;
     if PERSIST1 lt &P133_3 & PERSIST1 ne . then call define(_col_,"style","style={background=#fff7bc}");
	 else if PERSIST1 gt &P166_6 & PERSIST1 ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if PERSIST1 ne . then call define(_col_,"style","style={background=#bdbdbd}");
	 if find(_c1_,'All','i') then  CALL DEFINE("_c3_", "style","STYLE=[BACKGROUND=#ffffff]");
	
 endcomp;
 compute PERSIST2;
     if PERSIST2 lt &P233_3 & PERSIST2 ne . then call define(_col_,"style","style={background=#fff7bc}");
	 else if PERSIST2 gt &P266_6 & PERSIST2 ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if PERSIST2 ne . then call define(_col_,"style","style={background=#bdbdbd}");
	 if find(_c1_,'All','i') then  CALL DEFINE("_c4_", "style","STYLE=[BACKGROUND=#ffffff]");
	
 endcomp;
 compute PERSIST3;
     if PERSIST3 lt &P333_3 & PERSIST3 ne . then call define(_col_,"style","style={background=#fff7bc}");
	 else if PERSIST3 gt &P366_6 & PERSIST3 ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if PERSIST3 ne . then call define(_col_,"style","style={background=#bdbdbd}");
	 if find(_c1_,'All','i') then  CALL DEFINE("_c5_", "style","STYLE=[BACKGROUND=#ffffff]");
	
 endcomp;
compute GRAD4;
     if GRAD4 lt &G433_3 & GRAD4 ne . then call define(_col_,"style","style={background=#fff7bc}");
	 else if GRAD4 gt &G466_6 & GRAD4 ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if GRAD4 ne . then call define(_col_,"style","style={background=#bdbdbd}");
	 if find(_c1_,'All','i') then  CALL DEFINE("_c6_", "style","STYLE=[BACKGROUND=#ffffff]");
	
 endcomp;
 compute GRAD5;
     if GRAD5 lt &G533_3 & GRAD5 ne . then call define(_col_,"style","style={background=#fff7bc}");
	 else if GRAD5 gt &G566_6 & GRAD5 ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if GRAD5 ne . then call define(_col_,"style","style={background=#bdbdbd}");
	 if find(_c1_,'All','i') then  CALL DEFINE("_c7_", "style","STYLE=[BACKGROUND=#ffffff]");
	
 endcomp;
  compute GRAD6;
     if GRAD6 lt &G633_3 & GRAD6 ne . then call define(_col_,"style","style={background=#fff7bc}");
	 else if GRAD6 gt &G666_6 & GRAD6 ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if GRAD6 ne . then call define(_col_,"style","style={background=#bdbdbd}");
	 if find(_c1_,'All','i') then  CALL DEFINE("_c8_", "style","STYLE=[BACKGROUND=#ffffff]");
	
 endcomp;

compute Dept_1st;
     if find(Dept_1st,'All','i') then call define(_row_,"style","style={font_weight=bold}");
	
 endcomp;
 run;

 /**PAG graduating**/
 data PAGdg;
set PAG_degr;
MAU= substr(Coll_Degr,1,2);
if MAU="&&coll&k";
call symput('degrcoll',Coll_Degr);
run;


proc report data= pagdg  split='~'  spanrows  style(header)={font_face='Arial, Helvetica'};
title1 h=13pt "First-Time Undergraduates Time-To-Degree By Graduating Cohort and Degree Department" ;
title2 h=6pt "";
title3 h=12pt &degrcoll ;
title4 "Graduating Cohort :" &gradcohort ;
footnote1 h=9pt justify=left "Data Source: PAG " ;
footnote2 h=9pt justify=left " ";
column Dept_DEGR N TTD;
define Dept_Degr/display 'Degree Department'  style(column)=[vjust=m just=l width=50%] ;
define N/ display 'N'  format=comma10.0;
define TTD/ display 'Time To Degree in Years' format=comma10.2  style(column)=[ width=20%];

compute TTD;
     if TTD lt &TTD33_3 & TTD ne . then call define(_col_,"style","style={background=#a1d99b }");
	 else if TTD gt &TTD66_6 & TTD ne . then call define(_col_,"style","style={background=#fff7bc}");
	 else if TTD ne . then call define(_col_,"style","style={background=#bdbdbd}");
	 if find(_c1_,'All','i') then  CALL DEFINE("_c3_", "style","STYLE=[BACKGROUND=#ffffff]");
	
 endcomp;
 
compute Dept_DEGR;
     if find(Dept_DEGR,'All','i') then call define(_row_,"style","style={font_weight=bold}");
	
 endcomp;
 run;
%end;
%mend ppsmetric;



options   nodate nonumber leftmargin=0.5in rightmargin=0.5in orientation=landscape papersize=A4 center missing=' ' ;
ods listing close;
ods pdf file="O:\IS\Internal\Reporting\Annual Reports\Key Metrics Set\Data\KeyMetrics_UngrdDept.pdf" style=Custom;
%ppsmetric;
ods pdf close;





