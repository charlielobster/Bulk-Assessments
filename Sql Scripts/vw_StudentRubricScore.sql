create view vw_StudentRubricScores 
as 
select 
	LabPK,
	sr.FKStudentPK,
	hcs.StudentReportPK,
	RubricGroupPK,
	round(cast(sum(case when hcs.HighConfidence = 1 then hcs.Score else 0 end) as float) / count(*), 2) [Score]
from 
(
	select 
		count(*) [Count], 
		sum(HighConfidence) [Sum HC],  
		sum(Score) [Sum Score], 
		round(cast(sum(HighConfidence) as float) / count(*), 2) [% HC],
		round(cast(sum(Score) as float) / count(*), 2) [% Score],
		LabPK,
		RubricGroupPK,
		RubricPK,
		RuleName,
		RuleLaTeX,
		RulePK 
	from vw_HighConfidenceScores 
	join vw_LabRuleDetails lrd
	on FKRulePK = RulePK
	group by 	
		LabPK,
		RubricGroupPK,
		RubricPK,
		RuleName,
		RuleLaTeX,
		RulePK 
	-- order by [Sum Score]
) t
join vw_HighConfidenceScores hcs
on hcs.FKRulePK = t.RulePK and hcs.FKLabPK = t.LabPK
join StudentReport sr
on hcs.StudentReportPK = sr.PK 
join Rubric r
on t.RubricPK = r.PK
where .25 <= [% Score] and [% HC] > .67 -- Remove 'Developing' and 'Incomplete' Rules
group by LabPK, sr.FKStudentPK, hcs.StudentReportPK, RubricGroupPK

go


select * from vw_StudentRubricScores

select	
	FKStudentPK,
	round([1] * 100, 2),
	round([2] * 100, 2),
	round([3] * 100, 2),
	round([4] * 100, 2)
from 
(
	select FKStudentPK, LabPK, Score from vw_StudentRubricScores
	where RubricGroupPK = 3
) src
pivot
(
	sum(Score)
	for LabPK in ([1], [2], [3], [4], [5])
) piv

select 	
	FKStudentPK,
	round([1] * 100, 2),
	round([2] * 100, 2),
	round([3] * 100, 2),
	round([4] * 100, 2)
from 
(
	select FKStudentPK, LabPK, Score from vw_StudentRubricScores
	where RubricGroupPK = 2
) src
pivot
(
  sum(Score)
  for LabPK in ([1], [2], [3], [4], [5])
) piv

select 	
	FKStudentPK,
	round([1] * 100, 2),
	round([2] * 100, 2),
	round([3] * 100, 2),
	round([4] * 100, 2)
from 
(
	select FKStudentPK, LabPK, Score from vw_StudentRubricScores
	where RubricGroupPK = 1
) src
pivot
(
  sum(Score)
  for LabPK in ([1], [2], [3], [4], [5])
) piv





drop view vw_StudentRubricScores
go