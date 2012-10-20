--Example 35-2. A PL/pgSQL Trigger Procedure

--This example trigger ensures that any time a row is inserted or updated in the table, 
--the current user name and time are stamped into the row. And it checks that an employee's name is
--given and that the salary is a positive value. 

CREATE FUNCTION rcljobs_stamp() RETURNS trigger AS $clients_stamp$
  DECLARE
        jobno           integer;
        clientno      integer;
       ojobno           integer;
   BEGIN
        IF (TG_OP = 'DELETE') THEN

            ojobno = cast (clients.leads_no as int4);
            jobno = cast (rcljobs.job_no as int4);
            clientno = cast (rcljobs.client_no as int4);
            if (jobno =  ojobno) THEN
                 jobno=jobno-1;
                 update clients set leads_no=cast(jobno as text)
                 where client_no=cast(clientno as text);
            end if;
       
       end if;
    END;
$clients_stamp$ LANGUAGE plpgsql;

CREATE TRIGGER clients_stamp AFTER DELETE ON rcljobs
    FOR EACH ROW EXECUTE PROCEDURE rcljobs_stamp();