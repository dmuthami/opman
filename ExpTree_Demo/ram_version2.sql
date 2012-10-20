 
--create table contact (
--	client_no      character varying,
--	description 	text,
--	f_name 			character varying,
	--s_name 			character varying,
	--salutation 		character varying,
	--pobox 			text,
	--e_mail1 		text,
	--e_mail2 		text,
--	fax             character varying,
--	tel				character varying,
--	cell			character varying,
--	physicaladd     character varying	
--		      );
--alter table clients add column least_status varchar;
--create table leads (
--	leads_no 	varchar,
--	client_no 	varchar,
--	descrip		varchar,
--	status		varchar
--		    );
		    
--alter table leads add column date_sniffed timestamp;
--alter table clients add column leads_no varchar;
--alter table clients add column oclient_no varchar;
--alter table rcljobs add column ojob_no varchar;
--alter table leads add column title varchar;
--alter table leads add column journal varchar;
--alter table leads
--alter column journal type text;
--alter table leads add column title varchar;
--alter table leads add column amount varchar;
--alter table rcljobs add column descrip varchar;

--changing the jobs
--alter table rcljobs add column date_sniffed timestamp;
--alter table rcljobs add column amount varchar;
--alter table rcljobs add column journal varchar;

---/////for version 2.3 only
   --personnel information
create table personnel_info (
	namme					 text,
	id_no 						character varying,
	hourly_rate 			character varying,
	gender 					character varying
		      );
    -- add an id_no to the security check table
    alter table seccheck add column id_no varchar;
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
   alter table archive_time add column milliseconds varchar;
	-- rcljobs_expenses	      
		      create table rcljobs_expenses	  (
	accomodation     	character varying,
	job_no 					character varying,
	travel					character varying,
	description 			text,
	amount				character varying
	
		      );	 
-- altering personnel_info table
alter table personnel_info  add column phone_no varchar;
alter table personnel_info add column mobile_no varchar;
alter table personnel_info add column postal_address text;
alter table personnel_info add column email text;
alter table personnel_info add column pin_no varchar;

--alering seccheck table
insert into seccheck (name,password,seclevel) values (
'Administrator','admin','1,1,1,1,1');

update seccheck set id_no='22040031' where name='Njue';
update seccheck set id_no='A636250' where name='DHWH';
update seccheck set id_no='20111801' where name='wns';
update seccheck set id_no='12955892' where name='LEONARD';
update seccheck set id_no='21501480' where name='MOHAMMED';	
update seccheck set id_no='13505221' where name='RAMSEY';
update seccheck set id_no='14427325' where name='MARTIN';
update seccheck set id_no='21759600' where name='IRENE';
insert into seccheck (name,password,id_no) values (
'Joshua','joshua','13349144');
      
update seccheck set id_no='20137594' where name='LINET';
update seccheck set id_no='22073499' where name='CHARLES';
update seccheck set id_no='21701213' where name='WYCLIFFE';
update seccheck set id_no='10861607' where name='Mutinda';		      
update seccheck set id_no='21236628' where name='REGINA';
update seccheck set id_no='21199927' where name='Rotich';
insert into seccheck (name,password,id_no) values (
'Alex','Alex','22171926');
insert into seccheck (name,password,id_no) values (
'Annastacia','Annastacia','21086326');

update seccheck set id_no='1' where name='JOHN';
update seccheck set id_no='2' where name='WILLIAM';
update seccheck set id_no='3' where name='Kebaso';


 insert into personnel_info (namme,id_no) values (
'Administrator','221345678');
insert into personnel_info (namme,id_no) values (
'njue','22040031');
 insert into personnel_info (namme,id_no) values (
'DHWH','A636250');
 insert into personnel_info (namme,id_no) values (
'wns','20111801');
 insert into personnel_info (namme,id_no) values (
'LEONARD','12955892');
 insert into personnel_info (namme,id_no) values (
'MOHAMMED','21501480');
 insert into personnel_info (namme,id_no) values (
'RAMSEY','13505221');
 insert into personnel_info (namme,id_no) values (
'MARTIN','14427325');
 insert into personnel_info (namme,id_no) values (
'IRENE','21759600');
insert into personnel_info (namme,id_no) values (
'Joshua','13349144');
      
 insert into personnel_info (namme,id_no) values (
'LINET','20137594');
 insert into personnel_info (namme,id_no) values (
'CHARLES','22073499');
 insert into personnel_info (namme,id_no) values (
'WYCLIFFE','21701213');
 insert into personnel_info (namme,id_no) values (
'Mutinda','10861607');
 insert into personnel_info (namme,id_no) values (
'REGINA','21236628');
 insert into personnel_info (namme,id_no) values (
'Rotich','21199927');  

insert into personnel_info (namme,id_no) values (
'Alex','22171926');
insert into personnel_info (namme,id_no) values (
'Annastacia','21086326'); 
insert into personnel_info (namme,id_no) values (
'JOHN','1');
insert into personnel_info (namme,id_no) values (
'WILLIAM','2'); 
   insert into personnel_info (namme,id_no) values (
'Kebaso','3'); 
   
 --equipments ----------------
 		      create table equip_info (
	equip_id					character varying,
	manufacturer 		character varying,
	model_no					character varying,
	serial_no 				character varying,
	model_name			character varying,
	purchase_date		timestamp,
	description 				text,
	license					character varying,
	guarantee				character varying,
	condition					character varying,
	type							character varying,
	model_year 				character varying
		      );
 --current equipments
 	      create table current_equip (
	equip_id								character varying,
	job_no									character varying,
	other									character varying,
	task										character varying,
	description 							text,
	assigned_by							character varying,
	date_assigned						timestamp,
	estimate_release_date		timestamp
	
		      );
   alter table current_equip add column date_released	 timestamp  ;
   alter table current_equip add column autonumber	 character varying  ;
 --historical data for equipments
       create table history_equip (
	equip_id								character varying,
	job_no									character varying,
	other									character varying,
	task										character varying,
	description 							text,
	assigned_by							character varying,
	date_assigned						timestamp,
	estimate_release_date		timestamp
	
		      );
 alter table history_equip add column autonumber	 character varying  ;
		 
alter table history_equip add column date_released	 timestamp  ;
     ALTER TABLE history_equip
		ALTER COLUMN date_assigned TYPE character varying,
		ALTER COLUMN estimate_release_date TYPE character varying,
		ALTER COLUMN date_released TYPE character varying;
		
		ALTER TABLE current_equip
		ALTER COLUMN date_assigned TYPE character varying,
		ALTER COLUMN estimate_release_date TYPE character varying,
		ALTER COLUMN date_released TYPE character varying;
		      --assignment info for equipments
		create table assigned_info (
	equip_id							character varying,
	status								character varying
		      );
		       create table equip_finances (
	equip_id								character varying,
	hourly_rate							character varying
	
		      );
		         create table maintenance_info (
	equip_id								character varying,
	service_date						timestamp,
	description							text,
	cost_incurred						character varying,
	invoice_no							character varying

		      );   
 alter table maintenance_info add column autonumber	 character varying  ; 
 
 create table tblno (
auto_no			character varying );   
 -- --------------------------end of equipments
 
 -----------------table for controlling time
        create table times (
	today							timestamp,
	usedate						character varying
	);
	
	
	--------------------version 2.4 
	--------leads
	  alter table leads add column department varchar;
	  alter table rcljobs add column department varchar;
	  
	  alter table contacts add column mobile1 varchar;
	  alter table contacts add column mobile2 varchar;
	  
	  -- altering personnel_info table
alter table personnel_info  add column birthday varchar;
alter table personnel_info add column contract_end text;
alter table personnel_info add column nssf_no varchar;
alter table personnel_info add column nhif_no varchar;
alter table personnel_info add column medical_cover text;

alter table personnel_info add column dateofemployment timestamp;
alter table personnel_info add column nextofkin varchar;
alter table personnel_info add column dateoftermination timestamp;
alter table personnel_info add column imagefile text;
--alter table personnel_info  drop column birthday ;
--alter table personnel_info drop column contract_end ;
--alter table personnel_info drop column nssf_no ;
--alter table personnel_info drop column nhif_no ;
--alter table personnel_info drop column medical_cover ;

--alter table personnel_info drop column dateofemployment ;
--alter table personnel_info drop column nextofkin ;
--alter table personnel_info drop column dateoftermination ;

 -----------------table for leaves
        create table leaves (
	idno						character varying,
	description	        text,
	sdate					timestamp,
	edate                   timestamp,
	ano                      serial
	);
	
	   create table sickoff (
	idno						character varying,
	description	        text,
	sdate					timestamp,
	edate                   timestamp,
	ano                      serial
	);
	
	   create table dayoff (
	idno						character varying,
	description	        text,
	dateoff					timestamp,
	timeoff                 time,
	ano                      serial
	);
	   create table timeoff (
	idno						character varying,
	description	        text,
	dateoff					timestamp,
	timeoff                 time,
	ano                      serial
	);
		   create table casuals (
	job_no					character varying,
	description	        text,
	task				        text,
	datehired            timestamp,
	wagespaid           character varying,
	ano                      serial
	);
	alter table casuals add column namme text;
		   create table accomodation (
	job_no					character varying,
	description	        text,
	costincurred       character varying,
	ano                      serial
	);
			   create table travel (
	job_no					character varying,
	description	        text,
	costincurred       character varying,
	ano                      serial
	);
	
	-- today is may 10 2006
	alter table rcljobs add column budgetarycost character varying;
	alter table rcljobs add column grossmargin character varying;
	
	alter table travel add column kilometers character varying;
	alter table travel add column othermodes character varying;
	alter table travel add column namme character varying;
	
	alter table accomodation add column namme character varying;
	alter table accomodation add column hotel character varying;
	alter table accomodation add column entry character varying;
	---------hired equipments
	
	create table hiredequip (
	equipname					character varying,
	description	                text,
	assigndate                  timestamp,
	releasedate                 timestamp,
	hourly_rate                 character varying,
	cost                              character varying,
	ano                              serial
	);
	
	alter table hiredequip add column job_no character varying;
	alter table equip_info add column hourly_rate character varying;
	
	--------today is may 11 2006
	
	alter table daily_time add column ano  serial;
	alter table archive_time add column ano  serial;
	
	
	create table grossmargin (
	personnel					  character varying,
	casual	                          character varying,
	accomodation               character varying,
	travel                            character varying,
	ramani                         character varying,
	hired                             character varying,
	job_no                            character varying,
	ano                                serial
	);
	
	
	
	--------------------------------------------post instal
	--alering seccheck table
update  seccheck  set id_no='221345678' where name='Administrator';
-------------






