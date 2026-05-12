create view vw_HCWeights as
select 
	t.RulePK,
	round(cast([Total HC] as float) / cast([Count] as float), 2) [HCWeight]
from (
	select 
		RulePK,
		count(*) [Count],
		isnull(sum(hcs.[HighConfidence]),0) [Total HC]
	from vw_LabRuleDetails lrd
	left join vw_HighConfidenceScores hcs
	on hcs.FKRulePK = lrd.RulePK and hcs.FKLabPK = lrd.LabPK
	where labPK != 6
	group by RulePK
	--order by RulePK
) t
join RubricRule rr on t.RulePK = rr.PK
join Rubric r on rr.FKRubricPK = r.PK
order by RulePK
go

drop view vw_HCWeights
go
