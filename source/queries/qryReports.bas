SELECT Outcomes.Outcome, Outcomes.Date, Outcomes.Notes
FROM Outcomes
WHERE (((Outcomes.Outcome) IN(Forms!Reporting!txtString)));

