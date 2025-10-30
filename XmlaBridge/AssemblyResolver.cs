using System;
using System.IO;
using System.Reflection;

public static class AssemblyResolver
{
    private static bool isInitialized = false;
    private static readonly object lockObject = new object();

    public static void Initialize()
    {
        if (isInitialized)
        {
            return;
        }

        lock (lockObject)
        {
            if (isInitialized)
            {
                return;
            }

            Console.WriteLine("[AssemblyResolver] Initializing assembly resolver...");
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
            isInitialized = true;
            Console.WriteLine("[AssemblyResolver] Assembly resolver initialized");
        }
    }

    private static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
    {
        try
        {
            var assemblyName = new AssemblyName(args.Name);
            Console.WriteLine($"[AssemblyResolver] Resolving: {assemblyName.Name}, Version={assemblyName.Version}");

            // Get the directory where XmlaBridge.dll is located
            string assemblyDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            // Try to find the assembly in the same directory
            string assemblyFileName = assemblyName.Name + ".dll";
            string assemblyPath = Path.Combine(assemblyDirectory, assemblyFileName);

            if (File.Exists(assemblyPath))
            {
                Console.WriteLine($"[AssemblyResolver] Loading from: {assemblyPath}");
                var assembly = Assembly.LoadFrom(assemblyPath);
                Console.WriteLine($"[AssemblyResolver] Loaded {assemblyName.Name} version: {assembly.GetName().Version}");
                return assembly;
            }
            else
            {
                Console.WriteLine($"[AssemblyResolver] Assembly not found at: {assemblyPath}");
            }

            return null;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[AssemblyResolver] Error resolving assembly: {ex.Message}");
            return null;
        }
    }
}
