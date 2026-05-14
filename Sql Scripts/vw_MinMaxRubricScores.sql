create view vw_MinMaxRubricScores
as 
	select 
		LabPK, 
		RubricGroupPK,
		max(Score) - min(Score) [Score Range],
		min(Score) [Min Score],
		max(Score) [Max Score],
		round(avg(Score), 2) [Avg Score]
		from vw_StudentRubricScores
		group by LabPK, RubricGroupPK

go

--drop view vw_MinMaxRubricScores
--go
