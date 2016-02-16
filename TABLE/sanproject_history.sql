 create table sanproject_history
 (
 user_id  VARCHAR2(15) CONSTRAINT userid_fk REFERENCES  hrishi.sanproject_user(user_id)
                                                        ON DELETE CASCADE,
 working_date    DATE DEFAULT SYSDATE,
 stime   VARCHAR2(15),
 Etime   VARCHAR2(15)
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
