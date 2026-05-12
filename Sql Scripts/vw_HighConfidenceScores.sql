create view vw_HighConfidenceScores as
select 
	sr.FKStudentPK, 
	sr.FKLabPK, 
	sr.PK [StudentReportPK],
	hcs.FKRulePK,
	hcs.HighConfidence,
	hcs.Score
from (
select 
	ScoresFilePK,
	FKRulePK,
	(case when [TotalScore] = 0 or [TotalScore] = 3 then 1 else 0 end) [HighConfidence],
	(case when [TotalScore] = 3 then 1 else 0 end) [Score] 
from (
	select 
		a.FKScoresFilePK ScoresFilePK, 
		rs.FKRulePK,
		sum(cast(rs.Score as int)) [TotalScore]
	from RuleScore rs join Assessment a on rs.FKAssessmentPK = a.PK
	group by a.FKScoresFilePK, rs.FKRulePK
	) t
) hcs
join ScoresFile sf on hcs.ScoresFilePK = sf.PK
join StudentReport sr on sf.FKStudentReportPK = sr.PK

go

select * from vw_HighConfidenceScores
order by FKStudentPK, FKLabPK, FKRulePK

--drop view vw_HighConfidenceScores
--go
