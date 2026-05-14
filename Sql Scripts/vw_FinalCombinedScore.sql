create view vw_FinalCombinedScore
as
select 
	LabPK,
	FKStudentPK,
	StudentReportPK,
	sum(case when RubricGroupPK = 1 then [Final Score] else 0 end) [Final Core Score],
	sum(case when RubricGroupPK = 2 then [Final Score] else 0 end) [Final Lab Score],
	sum(case when RubricGroupPK = 3 then [Final Score] else 0 end) [Final Bonus Score],
	( ( sum(case when RubricGroupPK = 1 then [Final Score] else 0 end)  +
	sum(case when RubricGroupPK = 2 then [Final Score] else 0 end) ) +
	sum(case when RubricGroupPK = 3 then [Final Score] else 0 end) * .5 ) / 2.5 [Final Combined Score]
from vw_RubricFinalScores
group by 
	LabPK,
	FKStudentPK,
	StudentReportPK

go

-- drop view vw_FinalCombinedScore
-- go
