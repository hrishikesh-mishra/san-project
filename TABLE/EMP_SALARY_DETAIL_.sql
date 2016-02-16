create table HRISHI.EMP_SALARY_DETAIL
(
eid  NUMBER(9) CONSTRAINT empid_nn NOT NULL,
ename VARCHAR2(25) CONSTRAINT empname_nn NOT NULL,
sal_for_month  CHAR(12) CONSTRAINT monthName_nn NOT NULL,
basic   NUMBER(10,2),
hra     NUMBER(2,2),
da      NUMBER(2,2),
ta      NUMBER(2,2),
deduction   NUMBER(2,2),
tax         NUMBER(2,2),
special_pay NUMBER(5,2),
festival_pay NUMBER(5,2),
gross_sal   NUMBER(12,2),
total_deduction NUMBER(12,2),
net_sal   NUMBER(12,2),
entry_date  DATE DEFAULT SYSDATE
)
PCTFREE 5
PCTUSED 30
INITRANS 10
MAXTRANS 15
TABLESPACE SAN
STORAGE(INITIAL 50K
        NEXT 50K
        MAXEXTENTS 10
        PCTINCREASE 20)

/
