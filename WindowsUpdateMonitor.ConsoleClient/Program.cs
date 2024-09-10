using WUApiLib;

namespace WindowsUpdateMonitor.ConsoleClient;

public class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Windows Update Agent Demonstration");
        Console.WriteLine("-----------------------------------");

        // Create an instance of the Windows Update Session
        UpdateSession updateSession = new UpdateSession();
        IUpdateSearcher updateSearcher = updateSession.CreateUpdateSearcher();

        // Display information about available updates
        DisplayAvailableUpdates(updateSearcher);

        // Display information about installed updates
        DisplayInstalledUpdates(updateSearcher);

        // Display information about hidden updates
        DisplayHiddenUpdates(updateSearcher);

        Console.WriteLine("Finished displaying all update information.");
        Console.ReadKey();
    }

    // Method to display available updates
    private static void DisplayAvailableUpdates(IUpdateSearcher updateSearcher)
    {
        Console.WriteLine("\nChecking for Available Updates...");
        try
        {
            ISearchResult searchResult = updateSearcher.Search("IsInstalled=0");
            DisplayUpdateCollection("Available Updates", searchResult.Updates);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while searching for available updates: {ex.Message}");
        }
    }

    // Method to display installed updates
    private static void DisplayInstalledUpdates(IUpdateSearcher updateSearcher)
    {
        Console.WriteLine("\nChecking for Installed Updates...");
        try
        {
            ISearchResult searchResult = updateSearcher.Search("IsInstalled=1");
            DisplayUpdateCollection("Installed Updates", searchResult.Updates);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while searching for installed updates: {ex.Message}");
        }
    }

    // Method to display hidden updates
    private static void DisplayHiddenUpdates(IUpdateSearcher updateSearcher)
    {
        Console.WriteLine("\nChecking for Hidden Updates...");
        try
        {
            ISearchResult searchResult = updateSearcher.Search("IsHidden=1");
            DisplayUpdateCollection("Hidden Updates", searchResult.Updates);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while searching for hidden updates: {ex.Message}");
        }
    }

    // Method to display updates in a collection
    private static void DisplayUpdateCollection(string title, IUpdateCollection updates)
    {
        Console.WriteLine($"\n--- {title} ---");
        if (updates.Count == 0)
        {
            Console.WriteLine("No updates found.");
            return;
        }

        for (int i = 0; i < updates.Count; i++)
        {
            IUpdate update = updates[i];
            Console.WriteLine($"Update {i + 1}:");
            Console.WriteLine($"- Title: {update.Title}");
            Console.WriteLine($"- Description: {update.Description}");
            Console.WriteLine($"- KB Article IDs: {string.Join(", ", update.KBArticleIDs)}");
            Console.WriteLine($"- Categories: {GetCategories(update)}");
            Console.WriteLine($"- More Info URL: {update.MoreInfoUrls}");
            Console.WriteLine($"- Support URL: {update.SupportUrl}");
            Console.WriteLine($"- Is Mandatory: {update.IsMandatory}");
            Console.WriteLine($"- Is Downloaded: {update.IsDownloaded}");
            Console.WriteLine($"- Is Installed: {update.IsInstalled}");
            Console.WriteLine($"- Is Hidden: {update.IsHidden}");
            Console.WriteLine($"- Last Deployment Change Time: {update.LastDeploymentChangeTime}");
            Console.WriteLine();
        }
    }

    // Helper method to get update categories
    private static string GetCategories(IUpdate update)
    {
        var categoryNames = "";
        foreach (ICategory category in update.Categories)
        {
            categoryNames += category.Name + ", ";
        }

        return categoryNames.TrimEnd(',', ' ');
    }
}