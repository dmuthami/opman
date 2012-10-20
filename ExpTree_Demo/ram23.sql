--drop table daily_time;
--drop table archive_time;
-- daily time sheet	      
		      create table daily_time (
	id_no					character varying,
	job_no 					character varying,
	task 					text,
	description 			text,
	ddate 					timestamp,
	timespent 			character varying
		      );
    alter table daily_time add column milliseconds  varchar;
   -- archive  time sheet	      
		      create table archive_time (
	id_no					character varying,
	job_no 					character varying,
	task 					text,
	description 			text,
	ddate 					timestamp,
	timespent 			character varying
		      );


-----------failed completely
alter table daily_time add column ano  serial;
alter table archive_time add column ano  serial;
	
alter table clients add column ano  serial;
alter table contact add column ano  serial;
	
alter table current_equip add column ano  serial;
	
	
alter table equip_info add column ano  serial;
alter table equip_finances add column ano  serial;
	
alter table history_equip add column ano  serial;
alter table leads add column ano  serial;
	
alter table rcljobs add column ano  serial;
alter table rcljobs_expenses add column ano  serial;

 alter table seccheck add column ano  serial;
 
alter table maintenance_info add column ano  serial;
alter table personnel_info add column ano  serial;
  -----------------failure ends here
 ----------------equipments------------------
alter table equip_info add column amount  character varying;
alter table equip_info add column supplier  character varying;


---------------------may 23 2006 
alter table equip_info add column mouse  character varying;
alter table equip_info add column keyboard  character varying;
alter table equip_info add column monitor  character varying;


--------------------------------------may 25 okay
  -- archive  storedate      
		      create table storedate (
	curdate 				timestamp,
	use 			            character varying,
	ano                      serial
		      );
		      
		      alter table daily_time add column stime  character varying;
		      alter table daily_time add column etime character varying;
		      
		       alter table archive_time add column stime  character varying;
		      alter table archive_time add column etime character varying;
		      
		       alter table daily_time add column notes  text;
		      alter table archive_time add column notes text;
		  
-------------------------------------

------------------------may 26
alter table seccheck add column arrdate  text;


----------------------
 --------------not in ramani database
-----------------------------------may 29
drop table current_equip;
drop table history_equip;
CREATE TABLE current_equip (
    equip_id character varying,
    job_no character varying,
    other character varying,
    task character varying,
    description text,
    assigned_by character varying,
    date_assigned character varying,
    estimate_release_date character varying,
    autonumber character varying,
    date_released character varying,
    ano serial NOT NULL
);
CREATE TABLE history_equip (
    equip_id character varying,
    job_no character varying,
    other character varying,
    task character varying,
    description text,
    assigned_by character varying,
    date_assigned character varying,
    estimate_release_date character varying,
    autonumber character varying,
    date_released character varying,
    ano serial NOT NULL
);
 update assigned_info set status='0';
alter table equip_info add column monitor2  character varying;
alter table equip_info add column phone  character varying;


---------------------------------------may 31 2006
alter table equip_info add column batteries  character varying;
alter table equip_info add column downloadcables  character varying;
alter table equip_info add column unit  character varying;

-----------------5 june

CREATE TABLE it (
    id_no character varying,
    issues  text,
    comments text,
    solved boolean,
    report_date timestamp,
    ano serial NOT NULL
);

----------------8 june
alter table personnel_info add column comments  character varying;