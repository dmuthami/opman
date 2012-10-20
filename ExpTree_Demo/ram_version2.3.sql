----------for may 18 2006
alter table leads
  alter column date_sniffed   type timestamp; 
  --contact
  	--alter table contact add column mobile2 character varying;
   alter table seccheck alter column seclevel type text;
   insert into seccheck (name,password,id_no) values ('administrator','admin','9999999999');
   insert into personnel_info (namme,id_no) values ('administrator','9999999999');
   update seccheck set seclevel='1,1,1,1,1,1:1,1,1,1,1,1:1,1,1,1,1,1:1,1,1,1,1,1:1,1,1,1,1,1' where id_no='9999999999';
   
   update seccheck set seclevel='1,1,1,1,1,1:1,1,1,1,1,1:1,1,1,1,1,1:1,1,1,1,1,1:1,1,1,1,1,1' where id_no='21501480';
   update seccheck set seclevel='0,0,0,0,0,0:0,0,0,0,0,0:0,0,0,0,0,0:0,0,0,0,0,0:0,0,0,0,0,0' where id_no<>'221345678';
   
   insert into seccheck (name,id_no) values ('Administrator','221345678');
insert into seccheck (name,id_no) values ('njue','22040031');
 insert into seccheck (name,id_no) values ('DHWH','A636250');
 insert into seccheck (name,id_no) values ('wns','20111801');
 insert into seccheck (name,id_no) values ('LEONARD','12955892');
 insert into seccheck (name,id_no) values ('MOHAMMED','21501480');
 insert into seccheck (name,id_no) values ('RAMSEY','13505221');
 insert into seccheck (name,id_no) values ('MARTIN','14427325');
 insert into seccheck (name,id_no) values ('IRENE','21759600');
insert into seccheck (name,id_no) values ('Joshua','13349144');
      
 insert into seccheck (name,id_no) values ('LINET','20137594');
 insert into seccheck (name,id_no) values ('CHARLES','22073499');
 insert into seccheck (name,id_no) values ('WYCLIFFE','21701213');
 insert into seccheck (name,id_no) values ('Mutinda','10861607');
 insert into seccheck (name,id_no) values ('REGINA','21236628');
 insert into seccheck (name,id_no) values ('Rotich','21199927');  

insert into seccheck (name,id_no) values ('Alex','22171926');
insert into seccheck (name,id_no) values ('Annastacia','21086326'); 
insert into seccheck (name,id_no) values ('JOHN','1');
insert into seccheck (name,id_no) values ('WILLIAM','2'); 
insert into seccheck (name,id_no) values ('Kebaso','3'); 