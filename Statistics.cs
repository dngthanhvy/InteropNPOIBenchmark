namespace InteropNPOIBenchmark;

public static class Statistics
{
    public static double Mean(List<long> values)
    {
        return values.Average();
    }
    
    public static long Min(List<long> values)
    {
        return values.Min();
    }
    
    public static long Max(List<long> values)
    {
        return values.Max();
    }
}