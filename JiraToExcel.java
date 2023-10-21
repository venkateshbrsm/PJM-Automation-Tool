public static void traverseAndPrint(Map<String, Object> map) {
        for (Map.Entry<String, Object> entry : map.entrySet()) {
            String key = entry.getKey();
            Object value = entry.getValue();

            if (value instanceof Map) {
                System.out.println("Key: " + key + " (Map)");
                traverseAndPrint((Map<String, Object>) value); // Recurse for nested maps
            } else if (value instanceof List) {
                System.out.println("Key: " + key + " (List)");
                List<?> list = (List<?>) value;
                for (int i = 0; i < list.size(); i++) {
                    System.out.println("  - Index " + i + ": " + list.get(i));
                }
            } else {
                System.out.println("Key: " + key + " (Value): " + value);
            }
        }
    }
}





