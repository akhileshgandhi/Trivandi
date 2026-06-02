export interface IPromotedResult {
  id: string;
  keywordMatches: string[];
  title: string;
  description: string;
  url: string;
  category: string;
  type: string;
}

export class PromotedResultsService {
  private static promotedList: IPromotedResult[] = [
    {
      id: 'promo_hr_portal',
      keywordMatches: ['hr', 'policy', 'leave', 'vacation', 'hr policy', 'benefits', 'medical', 'insurance'],
      title: 'Trivandi Corporate HR Portal',
      description: 'Official corporate human resources gateway. Apply for leaves, view health benefits coverage, submit claims, and view global HR policies.',
      url: 'https://sharepoint.trivandi.com/sites/hr/SitePages/HR-Home.aspx',
      category: 'Human Resources',
      type: 'portal'
    },
    {
      id: 'promo_holiday_calendar',
      keywordMatches: ['holiday', 'calendar', 'holidays', 'leaves', 'off', 'offices', 'festivals', '2026'],
      title: 'Trivandi Global Holiday Calendar 2026',
      description: 'Corporate holiday calendar for all international Trivandi office locations. Plan your leaves and check regional off-days.',
      url: 'https://sharepoint.trivandi.com/sites/hr/Shared%20Documents/Trivandi-Holiday-Calendar-2026.pdf?web=1',
      category: 'Company Info',
      type: 'pdf'
    },
    {
      id: 'promo_brand_guidelines',
      keywordMatches: ['brand', 'logo', 'style', 'guide', 'styleguide', 'design', 'assets', 'powerpoint', 'templates'],
      title: 'Trivandi Brand Identity & Style Guidelines',
      description: 'Access official vector logos, standard typography packages, visual brand guidelines, and approved PowerPoint templates.',
      url: 'https://sharepoint.trivandi.com/sites/marketing/Assets/Brand-Identity-Guidelines.pdf?web=1',
      category: 'Marketing',
      type: 'styleguide'
    },
    {
      id: 'promo_partnership_agreements',
      keywordMatches: ['partnership', 'contract', 'agreement', 'partner', 'vendor', 'legal', 'collaboration'],
      title: 'Approved Partnership Agreement Templates',
      description: 'Legal-approved master templates for corporate partnership agreements, NDA agreements, and vendor contracts.',
      url: 'https://sharepoint.trivandi.com/sites/legal/Templates/Master-Partnership-Agreement-2026.docx?web=1',
      category: 'Legal',
      type: 'template'
    }
  ];

  /**
   * Matches a search query against keyword arrays and returns matches
   */
  public static getPromotedResults(query: string): IPromotedResult[] {
    const clean = query.trim().toLowerCase();
    if (!clean) return [];

    return this.promotedList.filter(item => {
      return item.keywordMatches.some(keyword => {
        // Direct matching or checks if query contains the keyword
        return clean.includes(keyword) || keyword.includes(clean);
      });
    });
  }
}
