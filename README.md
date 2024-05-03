* Patterns which won't work
    * Hello &name_id& abcdef &&otp&&& - & cannot be parsed
    * (s) -> this is excluded
    * ,._ -> 3 are excluded, would mean ,abc, or .abc. or _abc_ won't work
    *  Previous working regex - // String regex = "[^a-zA-Z0-9_,. ]+[a-zA-Z_ ]+[^a-zA-Z0-9_,. ]+";

* Dependencies
  ```xml
  <dependency>
  <groupId>org.apache.poi</groupId>
  <artifactId>poi</artifactId>
  <version>5.2.5</version>
  </dependency>
  <dependency>
  <groupId>org.apache.poi</groupId>
  <artifactId>poi-ooxml</artifactId>
  <version>5.2.5</version>
  </dependency>
 
