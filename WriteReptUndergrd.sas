/***Connect to db***/
%include "H:\SAS\SAS_Log_EDW.sas" ;
LIBNAME BIMSUTST  odbc required="DSN=BIMSUTST; uid=&BIMSUTST_uid; pwd=&BIMSUTST_pwd";
libname PAG oracle path="MSUEDW" user="&MSUEDW_uid" pw= "&MSUEDW_pwd"
schema = OPB_PERS_FALL preserve_tab_names = yes connection=sharedread;
libname Devdm odbc required="DSN=NT_Dev_Datamart";
libname PPS odbc required="DSN=NT_PPS";

%let todaysDate = %sysfunc(today(), yymmddn8.);
%let AAver='AAU_112018';

option symbolgen;
options mlogic;

/*glossary data*/
proc import out=glossory datafile='O:\IS\Internal\Reporting\Annual Reports\Key Metrics Set\Key Metrics Glosary_V3.xlsx' dbms=xlsx replace; 
getnames=yes;
range="Sheet1$A3:C35";
run;

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
 quit;


/*set whether pps section include only departments have undergrads*/
%let ungradinclude=N;
/**whether include only depts with students versus all dept*/
%let alldept=Y;
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
                   375,81);

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

quit;




data pps;
set ppscoll ppsdp;
run;

/*transpose and compute for ratio**/
proc sort data= pps; by year_id year_name leadcoll Dept; run;
proc transpose data=pps out=ppsl prefix=V;
 by year_id year_name leadcoll Dept;
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
drop _NAME_;
run;
/**transpose back to long**/
proc transpose data=ppsl out =pps;
by year_id year_name leadcoll Dept;
run;
data pps;
set pps;
/**filter null value and numerator deminoaor */
where COL1 ne . and _NAME_ not in ('V37','V153','V375','V81','V108','V882','V883');
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
                   375,81);
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
drop _NAME_;
run;

/**transpose back to long**/
proc transpose data=yrlong1 out =yrlong1;
run;


proc sql;
create table pps as 
select a.*
from pps a
inner join yrlong1 b
on a._NAME_=b._NAME_ and a.year_id=b.COL1;
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
else var=_NAME_;
/*include only instruction college*/
where substr(leadcoll,1,2) in ('38', '08','10','02','14','32','16','04','24','46','28','33','34','22','30','06');
run;

/**transpose to wide 2 variables******/

proc sort data=pps; by leadcoll dept ;run;
/*for actual value*/
proc transpose data=pps out=pps1(drop=_NAME_ ) ;
id  var;
by leadcoll dept;
var COL1;
run;
/*for year*/
proc transpose data=pps out=pps2(drop=_NAME_ _LABEL_) prefix=Yr_;
id  var;
by leadcoll dept;
var year_name;
run;

data pps;
merge pps1 pps2;
by leadcoll dept;
run;

/*
proc print data= pps;
where substr(leadcoll,1,2)='14';
run;*/

/*********PPS metrics import pre processing from R ppskeymetricsds.R*************/
/********PPS metrics always using the latest year available*********************/

/*proc import datafile="O:\IS\Internal\Reporting\Annual Reports\Key Metrics Set\Data\PPSmetrics.xlsx"  out=PPSt dbms=xlsx replace ;sheet='PPS';
run;
/**PPS quantitle*
proc import datafile="O:\IS\Internal\Reporting\Annual Reports\Key Metrics Set\Data\PPSmetrics.xlsx"  out=PPSqt dbms=xlsx replace ;sheet='PPSquantitle';
run;*/

/*college with grad prof students*/
data pps;
set pps;
MANAGE_LEVEL_3= substr(leadcoll,1,2);
if find(Dept,'All','i')>0 then orderdp=1;
else orderdp=0;
run;

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
	
	/**include only those depts with students
where a.Total_Student ne . and a.Total_Student>0*/
order by a.MANAGE_LEVEL_3,a.orderdp,a.Dept;
quit;



%macro ppsds;
	%if &ungradinclude=Y %then %do;
	data pps;
	set pps;
	where Undergraduates ne . and Undergraduates>0 and Total_Student ne . and Total_Student>0;
	%end;
	%if &alldept=N %then %do;
	data pps;
	set pps;
	where  Total_Student ne . and Total_Student>0;
	%end;
	run;
%mend;
%ppsds;





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
 call symput('Ungrdyr',Yr_Undergraduates);
 call symput('Masteryr',Yr_Masters);
 call symput('Docyr',Yr_Doctoral);
 call symput('Profyr',Yr_Grad_Prof);
 call symput('V1yr',Yr_V1);
  call symput('V2yr',Yr_V2);
   call symput('V3yr',Yr_V3);
  call symput('V4yr',Yr_V4);
 call symput('V5yr',Yr_V5);
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
and Graduating_Cohort= "&gradcohort"
group by GRADUATING_COHORT, COLLEGE_DEGREE, COLLEGE_DEGREE_NAME, DEPT_DEGREE, DEPT_DEGREE_NAME;

create table PAG_Grad_Coll as 
select GRADUATING_COHORT, strip(COLLEGE_DEGREE)||'-'|| strip(COLLEGE_DEGREE_NAME) as Coll_DEGR, strip(COLLEGE_DEGREE)||'-'|| strip(COLLEGE_DEGREE_NAME)||'-All' as Dept_DEGR,
       count(PID) as N,AVG(TTD_IN_YEARS) AS TTD
from PAG.PERSISTENCE_V
where Student_LEVEL='UN'
and (  ENTRANT_SUMMER_FALL='Y' or substr(ENTRY_TERM_CODE,1,1)='F' )
and LEVEL_ENTRY_STATUS='FRST'
and Graduating_Cohort= "&gradcohort"
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
proc sql;
create table tocfor as 
select MANAGE_LEVEL_3, 1 as sec
 from PPS
 union
 select MAU, 2 as sec
 from AAprog
 union 
 select substr(Coll_1st,1,2) as Coll, 3 as sec
 from PAG_enter
 union 
 select substr(Coll_degr,1,2) as Coll,4 as sec
 from PAG_Degr;
 quit;

 /**preparing for table of contents in proc document macro*/
 proc sort data=tocfor; by  MANAGE_LEVEL_3 sec;
 data tocfor;
  set tocfor;
  seccount + 1;
  by MANAGE_LEVEL_3;
  if first.MANAGE_LEVEL_3 then seccount = 1;
run;
data tocfor;
set tocfor;
counter+1;
run;

proc sql;
create table toc as 
select MANAGE_LEVEL_3, min(counter) as min, count(*) as cnt, min(sec) as minsec
from tocfor
group by MANAGE_LEVEL_3;

create table tocfor as
select  a.MANAGE_LEVEL_3,a.counter,min, cnt, minsec
from tocfor a
inner join toc b 
on a.MANAGE_LEVEL_3=b.MANAGE_LEVEL_3
where (a.seccount >1) or cnt=1;
quit;
/*excluding 43*/
data _null_;
set tocfor;
call symput('dpg'||left(_n_),counter);
call symput('rpg'||left(_n_),min);
where cnt>1;
run;

proc sql noprint;
select count(*) into :seccnt
from tocfor
where cnt>1;
quit;

/*TOC adjust pages macro*/
%macro pageadj;
%do i=1 %to &seccnt;
move Report#&&dpg&i.\Report#1 to report#&&rpg&i.;
 delete Report#&&dpg&i.;
%end;
%mend;

/*macro for loop over all colleges*/
proc sql noprint;
select distinct count(distinct MANAGE_LEVEL_3)
 into :collcnt
 from toc;
quit;

data _null_;
set toc;
call symput('coll'||left(_n_),MANAGE_LEVEL_3);
call symput('minsec'||left(_n_),minsec);
run;

/*proc sql stimer;
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
*/





/*******set cutomized report template******************/


proc template;
 define style Styles.Custom;
 parent = Styles.Printer;
 Style SystemTitle from SystemTitle /
Font = ("Arial, Helvetica", 6pt, bold );
Style SystemFooter from systemFooter /
Font = ("Arial, Helvetica", 6pt, bold );
 
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



/***PPS colorblock***/
  %macro colorblock(var=V, t=PSV, j=2,add=5);
	  compute &&var.&j;
     if &&var.&j lt &&&t.&j.33_3 & &&var.&j ne . then call define(_col_,"style","style={background=#fff7bc}");
	 else if &&var.&j gt &&&t.&j.66_6 & &&var.&j ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if &&var.&j ne . then call define(_col_,"style","style={background=#bdbdbd}");
	 if find(_c1_,'All','i') then  CALL DEFINE("_c%eval(&j+&add)_", "style","STYLE=[BACKGROUND=#ffffff]");
       endcomp;
   %mend;
/***AA colorblock**/
%macro AAcolor(var,j);
	  compute &var;
     	 if &var lt &&&j.33_3 & &var ne . then call define(_col_,"style","style={background= #fff7bc}");
	 else if &var gt &&&j.66_6 & &var ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if &var ne . then call define(_col_,"style","style={background=#bdbdbd}");
       endcomp;
 %mend;


%macro ppsmetric;
	%do k=1 %to &collcnt;
		data ds;
		set pps;
		where MANAGE_LEVEL_3="&&coll&k";
		call symput('collname',leadcoll);
		call symput('gradprofn',gradprofcnt);
		run;
%if &&minsec&k=1 %then %do;
ods proclabel="%scan(%quote(&collname.),2,%str(-))";
%end;
%else %do;
ods proclabel=' '; 
%end;
proc report data= ds  style(report)={outputwidth=100%} style(header)={font_face='Arial, Helvetica' font_size=8pt} contents="PPS Metrics" ;
		title1 h=13pt "PPS Metrics by Lead Organization Structure"  ;
		title2 h=6pt " ";
		title3 h=10pt "&collname" ;
		footnote1 h=9pt justify=left "Note: master student counts include non-degree seeking lifelong graduates and MSU Law lifelong students." ;
		footnote2 h=9pt justify=left "Departments included are those led by instructional colleges and have data in at least one PPS metrics.";
		footnote3 h=9pt justify=left "PPS version: &version";
		footnote4 j = r 'Page !{thispage} of !{lastpage}';
		column MANAGE_LEVEL_3 Dept Undergraduates Masters Doctoral Grad_Prof V1 V2 V3 V4 V5;
		define Dept/display 'Dept Name' ;
		define Undergraduates/display "Undergrads Fall %substr(&Ungrdyr,1,4)" format=comma15.0 ;
		define Masters/display "Masters Fall %substr(&Masteryr,1,4)" format=comma15.0 ;
		define Doctoral/display "Doctoral Fall %substr(&Docyr,1,4)" format=comma15.0 ;
		%if &gradprofn=0 %then %do;
		define Grad_Prof/display "Professional Students Fall %substr(&Profyr,1,4)" format=comma15.0 noprint;
		%end;
		%else %do;
        define Grad_Prof/display "Professional Students Fall %substr(&Profyr,1,4)" format=comma15.0 ;
		%end;
		define V1/display "General Fund Budget/FY Admin SCH &V1yr" format=dollar15.0 style(column)=[width=8%];
		define V2/display "Admin Based SCH per Rank Fac FY &V2yr" format=comma15.0 style(column)=[width=8%];
		define V3/display "Tuition Revenue/Instructional Cost &V3yr" format=percentn15.0 style(column)=[width=8%];
		define V4/display "Total Grants - 3 Year Average-PI-Real$ &V4yr" format=dollar15.0;
		define V5/display "Total Grant PI Based/TS AF-FTE &V5yr" format=dollar15.0 style(column)=[width=8%];
		define MANAGE_LEVEL_3/group noprint; 
		break before MANAGE_LEVEL_3 / contents="" page;

	

 compute V1;
     if V1 lt &PSV133_3 & V1 ne . then call define(_col_,"style","style={background=#a1d99b}");
	 else if V1 gt &PSV166_6 & V1 ne . then call define(_col_,"style","style={background=#fff7bc}");
	 else if V1 ne . then call define(_col_,"style","style={background=#bdbdbd}");
	 if find(_c1_,'All','i') then  CALL DEFINE("_c6_", "style","STYLE=[BACKGROUND=white]");
	
 endcomp;
 %colorblock(j=2);
 %colorblock(j=3);
 %colorblock(j=4);
 %colorblock(j=5);
 compute Dept;
     if find(Dept,'All','i') then call define(_row_,"style","style={font_weight=bold}");
	
 endcomp;
 
run;

/**AA**/
data aa;
set AAprog;
MAUName= MAU||'-'|| MAUName;
where MAU="&&coll&k";
call symput('MAUName',MAUName);
call symput('AAver', VERSION);
call symput('Yr',YEAR );
run;

%if &&minsec&k=2 %then %do;
ods proclabel="%scan(%quote(&collname.),2,%str(-))";
%end;
%else %do;
*ods noproctitle;
ods proclabel=' '; 
%end;
proc report data= aa  split='~'  spanrows style(report)={outputwidth=100%}  style(header)={font_face='Arial, Helvetica' font_size=8pt} contents="Academic Analytics";
title1 h=13pt "Academic Analytics Metrics PhD Program View &Yr" ;
title2 h=6pt " ";
title3 h=10pt &MAUName ;
footnote1 h=9pt justify=left "Data Source: Academic Analytics Program View" ;
footnote2 h=9pt justify=left "AA version: &AAver";
footnote3 justify=left " ";
footnote4 j = r 'Page !{thispage} of !{lastpage}';
column MAU PHD_PROGRAM FACULTY_IN_PROG NATIONAL_ACADEMY_MEMBERS PROG_IN_MULT_TAXONOMIES TAXONOMY 
	AAU_PROG_IN_DISCIPLINE1 PCT_SCHLR_RSCH_INDEX1  RANK_SCHLR_RSCH_INDEX1 PCT_JOURNAL_PUBS_PER_FACULTY1 PCT_CITATIONS_PER_FACULTY1
 PCT_CITATIONS_PER_PUBLICATION1 PCT_GRANTS_PER_FACULTY1 PCT_GRANT_DOLLARS_PER_FACULTY1 PCT_DOLLARS_PER_GRANT1 PCT_AWARDS_PER_FACULTY1;
define PHD_PROGRAM/ order 'PhD Program'  style(column)=[vjust=m just=l ] ;
define FACULTY_IN_PROG/ order 'N of Faculty' style(column)=[vjust=m   ];
define NATIONAL_ACADEMY_MEMBERS/ order 'N of National Academy Member' style(column)=[vjust=m   ];
define PROG_IN_MULT_TAXONOMIES/ order 'Program Mutiple Taxonomy' style(column)=[vjust=m  ];
define TAXONOMY/display   'Taxonomy' style(column)=[cellwidth=10%]  ;
define AAU_PROG_IN_DISCIPLINE1/display 'N of Program in discipline at AAU' ;
define PCT_SCHLR_RSCH_INDEX1/display 'FSPI Percentile' format=comma10.2 ;
define RANK_SCHLR_RSCH_INDEX1/display 'FPSI Rank' format=comma10.0 ;
define PCT_JOURNAL_PUBS_PER_FACULTY1/display 'Percentile on Journal Pubs per faculty' format=comma10.2  ;
define PCT_CITATIONS_PER_FACULTY1/display 'Percentile on Citations per faculty' format=comma10.2 ;
define PCT_CITATIONS_PER_PUBLICATION1/display 'Percentile on Citations per publication' format=comma10.2 ;
define PCT_GRANTS_PER_FACULTY1/display 'Percentile on Grants per faculty' format=comma10.2  ;
define PCT_GRANT_DOLLARS_PER_FACULTY1/display 'Percentile on Grant Dollars per faculty' format=comma10.2 ;
define PCT_DOLLARS_PER_GRANT1/display 'Percentile Dollars per Grant' format=comma10.2 ;
define PCT_AWARDS_PER_FACULTY1/display 'Percentile on Awards per faculty' format=comma10.2 ;
define MAU / group noprint; 
break before MAU / contents="" page;
%AAcolor(var=PCT_JOURNAL_PUBS_PER_FACULTY1,j=JP);
%AAcolor(var=PCT_CITATIONS_PER_FACULTY1,j=CF);
%AAcolor(var=PCT_CITATIONS_PER_PUBLICATION1, j=CP);
%AAcolor(var=PCT_GRANTS_PER_FACULTY1, j=GF);
%AAcolor(var=PCT_GRANT_DOLLARS_PER_FACULTY1, j=GDF);
%AAcolor(var=PCT_DOLLARS_PER_GRANT1,j=DG);
%AAcolor(var=PCT_AWARDS_PER_FACULTY1,j=AF);
run;

/**PAGenter cohort**/
data PAG;
set PAG_enter;
MAU= substr(Coll_1st,1,2);
if MAU="&&coll&k";
call symput('entercoll',Coll_1st);
run;

%if &&minsec&k=3 %then %do;
ods proclabel="%UPCASE(%scan(%quote(&entercoll.),2,%str(-)))";
%end;
%else %do;
ods proclabel=' '; 
%end;
*ods noproctitle;
proc report data= pag  split='~'  spanrows style(report)={outputwidth=100%} style(header)={font_face='Arial, Helvetica' font_size=8pt}  contents="Persistence & Graduation" ;
title1 h=13pt "First-Time Undergraduates Persistence At and Graduation Rate From MSU By First Department" ;
title2 h=6pt " ";
title3 h=11pt &entercoll ;
title4 h=10pt "Entering Cohort : &entrycohort ";
footnote1 h=9pt justify=left "Data Source: PAG " ;
footnote2 h=9pt justify=left " ";
footnote3 j = r 'Page !{thispage} of !{lastpage}';
column MAU Dept_1st N PERSIST1 PERSIST2 PERSIST3 GRAD4 GRAD5 GRAD6;
define Dept_1st/display 'Entering Department'  style(column)=[vjust=m just=l ] ;
define N/ display 'N'  format=comma10.0;
define PERSIST1/ display 'Persist 1st Returning Fall' format=comma10.1 ;
define PERSIST2/ display 'Persist 2nd Returning Fall' format=comma10.1 ;
define PERSIST3/ display 'Persist 3rd Returning Fall' format=comma10.1 ;
define Grad4/ display 'Graduated by 4th Yr' format=comma10.1 ;
define Grad5/ display 'Graduated by 5th Yr' format=comma10.1 ;
define Grad6/ display 'Graduated by 6th Yr' format=comma10.1 ;
define MAU/group noprint;
break before MAU / contents="" page;
%colorblock(var=PERSIST,t=P,j=1,add=2);
%colorblock(var=PERSIST, t=P, j=2,add=2);
%colorblock(var=PERSIST, t=P, j=3,add=2);
%colorblock(var=GRAD, t=G, j=4,add=2);
%colorblock(var=GRAD, t=G, j=5,add=2);
%colorblock(var=GRAD, t=G, j=6,add=2);

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

%if &&minsec&k=4 %then %do;
ods proclabel="TIME-TO-DEGREE - %UPCASE(%scan(%quote(&degrcoll.),2,%str(-)))";
%end;
%else %do;
ods proclabel=' ';
%end;
proc report data= pagdg  split='~'  spanrows  style(header)={font_face='Arial, Helvetica'} contents="Time-To-Degree";
title1 h=13pt "First-Time Undergraduates Time-To-Degree By Graduating Cohort and Degree Department" ;
title2 h=6pt " ";
title3 h=11pt &degrcoll ;
title4 h=10pt "Graduating Cohort :" &gradcohort ;
footnote1 h=9pt justify=left "Data Source: PAG " ;
footnote2 h=9pt justify=left " ";
footnote3 j = r 'Page !{thispage} of !{lastpage}';
column MAU Dept_DEGR N TTD;
define Dept_Degr/display 'Degree Department'  style(column)=[vjust=m just=l width=50%] ;
define N/ display 'N'  format=comma10.0;
define TTD/ display 'Time To Degree in Years' format=comma10.2  style(column)=[ width=20%];
define MAU/group noprint;
break before MAU/ contents="" page; 
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


/**output prep for proc document*/
ods document name=mydocrep(write);
%ppsmetric;
/*glossary*/
ods proclabel="GLOSSARY";
proc print data=glossory  noobs   split='~'  style(report)={outputwidth=100%} style(header)={font_face='Arial, Helvetica' font_size=8pt}  contents="" ;
title1 h=13pt "Key Metrics Glossary: Planning Profile Summary, Academic Analytics, and Persistence and Graduation Data" ;
footnote1 j=l h=9pt ITALIC "Key Metrics Color Coding:" ;
footnote2 j=l h=8pt   "Green - The department ranks within the top third of departments at MSU on this metric. ";
footnote3 j=l h=8pt  "Grey - The department ranks within the middle third of departments at MSU on this metric. ";
footnote4 j=l h=8pt  "Yellow - The department ranks within the bottom third of departments at MSU on this metric. ";
footnote5 j=l h=8pt "White - This metric is not ranked. ";
footnote6 j = r 'Page !{thispage} of !{lastpage}';
run;
ods document close;


proc document name=mydocrep ;
/*adjust toc*/
 %pageadj;
run;
options   nodate  nonumber leftmargin=0.5in rightmargin=0.5in orientation=landscape papersize=A4 center missing=' ' ;
ods listing close;
ods pdf file="O:\IS\Internal\Reporting\Annual Reports\Key Metrics Set\Output\KeyMetrics_alldept_&todaysDate..pdf" style=Custom contents=yes;
ods escapechar= '!';
 replay;
run;
ods listing;
ods pdf close;
quit; 

