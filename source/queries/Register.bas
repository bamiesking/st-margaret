SELECT [Service Users].[Enrolled courses].Value, [Service Users].Surname, [Service Users].Forename
FROM [Service Users]
WHERE ((([Service Users].[Enrolled courses].Value)=[Forms].[Register].[RegSelBox]));

