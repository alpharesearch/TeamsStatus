namespace TeamsStatus;
class Program
{
    static void Main(string[] args)
    {
        MicrosoftTeamsService myMicrosoftTeamsService = new MicrosoftTeamsService();
        Console.WriteLine(myMicrosoftTeamsService.GetCurrentStatus().ToString());
    }
}