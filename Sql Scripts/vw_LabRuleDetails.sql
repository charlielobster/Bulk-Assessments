create view vw_LabRuleDetails as 
select distinct
	l.PK [LabPK],
	l.LabName,
	r.FKRubricGroupPK [RubricGroupPK],	
	rg.RubricGroupName,
	r.PK [RubricPK],
	r.RubricName,
	rr.PK [RulePK],
	rr.RuleID,
	rr.RuleName,
	rr.RuleLaTeX
from LabRubric lr 
join Lab l on l.PK = lr.FKLabPK
join Rubric r on r.PK = lr.FKRubricPK
join RubricRule rr on rr.FKRubricPK = r.PK
join RubricGroup rg on r.FKRubricGroupPK = rg.PK

go

select * from vw_LabRuleDetails
where LabName = 'Measurements'
order by LabPK, RubricGroupPK, RulePK

--drop view vw_LabRuleDetails
--go
