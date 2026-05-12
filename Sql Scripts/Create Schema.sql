-- create schema

create table RubricGroup
(
	PK int identity(1,1) primary key,
	RubricGroupName varchar(10) not null,
)

create table Rubric
(
	PK int identity(1,1) primary key,
	FKRubricGroupPK int foreign key references RubricGroup(PK) not null,
	Prefix varchar(2) not null unique,
	RubricName varchar(50) not null unique,
	RubricLaTeX varchar(max) not null,
)

create table RubricRule 
(
	PK int identity(1,1) primary key,
	FKRubricPK int foreign key references Rubric(PK) not null,
	RuleID varchar(5) not null unique,
	RuleName varchar(250) null,
	RuleLaTeX varchar(max) null,
)

create table Lab
(
	PK int identity(1,1) primary key,
	LabName varchar(50) not null unique,
)

create table LegacyFile
(
	PK int identity(1,1) primary key,
	FKLabPK int foreign key references Lab(PK) not null,
	LegacyFileName varchar(50) not null,
	LegacyFileBytes varbinary(max) not null,
)

create table LaTeXFile
(
	PK int identity(1,1) primary key,
	LaTeXFileName varchar(50) not null,
	LaTeX varchar(max) not null,	
)

create table LabRubric
(
	PK int identity(1,1) primary key,
	FKLabPK int foreign key references Lab(PK) not null,
	FKRubricPK int foreign key references Rubric(PK) not null,
)

create table Student
(
	PK int identity(1,1) primary key,
	StudentName varchar(300) not null unique,
	AnonymousID varchar(50) not null unique,
)

create table StudentReport
(
	PK int identity(1,1) primary key,
	FKStudentPK int foreign key references Student(PK) not null,
	FKLabPK int foreign key references Lab(PK) not null,
	ReportFileName varchar(50) not null,
	FileBytes varbinary(max) not null,
)

create table ScoresFile
(
	PK int identity(1,1) primary key,
	FKStudentReportPK int foreign key references StudentReport(PK) not null,
	ScoresFileName varchar(50) not null,
	ScoresFileBytes varbinary(max) not null,
)

create table Assessment
(
	PK int identity(1,1) primary key,
	FKScoresFilePK int references ScoresFile(PK) not null,
)

create table RuleScore
(
	PK int identity(1,1) primary key,
	FKRulePK int foreign key references RubricRule(PK) not null,
	FKAssessmentPK int foreign key references Assessment(PK) not null,
	Evidence varchar(max) not null,
	Score bit default(0),
)

-- insert some defaults
insert into RubricGroup select RubricGroupName from (values ('Core'), ('Lab'), ('Bonus')) as t(RubricGroupName)

insert into Student select StudentName, AnonymousID from 
(values 
	('Al Amin Olowu','AO'), 
	('Antonio Ramirez','AR'), 
	('Cindy Sookram', 'CS'), 
	('Dorian Guzman-Lora', 'DG'), 
	('Durjoy Roy', 'DR'), 
	('Daimon Son', 'DS'),
	('Elias Aguirre Carreno', 'EC'),
	('Farhan Chowdhury', 'FC'),
	('Fuad Rabbi', 'FR'),
	('Kervens Jean-Pierre', 'JK'),
	('Justyn Lopez Ullauri', 'JL'),
	('Jonncarlo Sagbaicela', 'JS'),
	('Kareem Alchorbaji', 'KA'),
	('Kamronbek Kalimbetov', 'KK'),
	('Mohamed Aboseria', 'MA'), 
	('Maysara Rahaman', 'MR'),
	('Paz Wang', 'PW'),
	('Rafael Santana', 'RS'),
	('Raphael Spoerri', 'RT'),
	('Shyam Persaud', 'SP'),
	('Sudhir Sukhai', 'SS'),
	('Tajbiul Mahmood', 'TM'),
	('William Gonzalez', 'WG'),
	('Wei Ling Thoo', 'WT'),
	('Yusuf Ahmed', 'YA')
) as t(StudentName, AnonymousID)

-- Clean-Up
 --delete from RuleScore
 --drop table RuleScore
 --delete from Assessment
 --drop table Assessment
 --delete from ScoresFile
 --drop table ScoresFile
 --delete from StudentReport
 --drop table StudentReport
 --delete from Student
 --drop table Student
 --delete from LabRubric
 --drop table LabRubric
 --delete from LegacyFile
 --drop table LegacyFile
 --delete from LaTeXFile
 --drop table LaTeXFile
 --delete from Lab
 --drop table Lab
 --delete from RubricRule
 --drop table RubricRule
 --delete from Rubric
 --drop table Rubric
 --delete from RubricGroup
 --drop table RubricGroup

