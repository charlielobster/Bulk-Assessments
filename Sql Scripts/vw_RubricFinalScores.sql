create view vw_RubricFinalScores as
select 
	srs.LabPK, 
	srs.FKStudentPK, 
	srs.StudentReportPK, 
	srs.RubricGroupPK, 
	Score, 
	round(( Score - [Min Score] ) / [Score Range], 2) [Normal Score],			-- normalize
	round(( Score - [Min Score] ) / [Score Range], 2) * 35 + 65 [Final Score]	-- reset the range to between 65-100
from vw_StudentRubricScores srs
join vw_MinMaxRubricScores mmrs on srs.LabPK = mmrs.LabPK
and srs.RubricGroupPK = mmrs.RubricGroupPK

go

-- drop view vw_RubricFinalScores
-- go