export interface ISearchFilters {
  fileType?: string;       // "pdf" | "docx" | "xlsx" | "pptx" | "png" | ""
  author?: string;
  modifiedDate?: string;   // "today" | "yesterday" | "week" | "month" | "year" | specific date
  specificDate?: string;   // ISO date if modifiedDate === "specific"
  site?: string;
}
