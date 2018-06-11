Key Metric Reporting
================

Summary
-------

The reports consists of :

1.  PPS metrics by lead college and department- only departments from instructional college and with undergrad components are included
2.  AA metrics by college and PHd Program
3.  PAG persistence and graduation rate by entering cohort
4.  PAG Time to Degree by graduating cohort

Data source
-----------

1.  Devdm.rowdata\_no\_extras
2.  AA Phd Program data from Kyle and uploaded to EDW with same format.BIMSUTST.ACADEMIC\_ANALYTICS\_METRICS\_P
3.  PAG OPB\_PERS\_FALL.PERSISTENCE\_V

STEP
----

SAS PROC REPORT is used for generating reports as PDF and for the conditional formatting on the cells
