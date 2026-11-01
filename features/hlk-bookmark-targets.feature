Feature: Check internal hyperlink bookmark targets
  In order to find internal hyperlinks whose bookmark target is missing or ambiguous
   As a developer preparing to generate a report from a document
  I need a read-only check that matches internal hyperlink anchors against bookmarks


  Scenario: Unique, missing, and duplicate bookmark targets are all reported
    Given a document containing bookmarks and internal hyperlinks
    When I check the internal hyperlink bookmark targets
    Then the report lists each internal link in document order with its text, anchor, and status
     And the duplicate-target result gives the number of matching bookmarks
     And external links, anchorless links, and field-code links are not reported


  Scenario: Links in nested and merged table cells are counted by physical XML occurrence
    Given a document containing internal hyperlinks in nested and merged table cells
    When I check the internal hyperlink bookmark targets
    Then each physical table cell contributes its links exactly once in document order


  Scenario: The check is read-only and results survive saving and reloading
    Given a checked document containing bookmarks and internal hyperlinks
    When I save and reload the document and check it again
    Then the reloaded document reports the same bookmark-target results
     And the hyperlinks, bookmarks, and relationships are unchanged by the checks
