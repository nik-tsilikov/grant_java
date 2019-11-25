public class spellchecker_levenshtein {
    public static void main(String[] args) {
        DamerauLevensteinMetric damerauLevensteinMetric = new DamerauLevensteinMetric();
        String string1 = "Доска";
        String string2 = "Даска";
        System.out.println(damerauLevensteinMetric.getDistance(string1, string2, 0));
    }
}
