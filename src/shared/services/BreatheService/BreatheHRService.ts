
export interface IBreatheHRDepartment {
    id: number;
    name: string;
}

export interface IBreatheHRLocation {
    id: number;
    name: string;
}

export interface IBreatheHRDivision {
    id: number;
    name: string;
}

export interface IBreatheHRLineManager {
    id: number;
    first_name: string;
    last_name: string;
    email: string;
}

export interface IBreatheHREmployee {
    id: number;
    first_name: string;
    last_name: string;
    known_as?: string;
    email: string;
    job_title?: string;
    status: string;
    photo_url?: string;
    work_mobile?: string;
    personal_mobile?: string;
    dob?: string;
    join_date?: string;
    department?: IBreatheHRDepartment;
    location?: IBreatheHRLocation;
    division?: IBreatheHRDivision;
    line_manager?: IBreatheHRLineManager;
    employee_ref?: string;
    contract_type?: string;
    full_or_part_time?: string;
    key_responsibility?: string;
    [key: string]: any;
}

const BREATHEHR_BASE_URL = "https://api.breathehr.com/v1";
const BREATHEHR_API_KEY = "prod-r0rkL1-d4hPDYhT6uPo_d5dtuQ3EuzzIDNTjxBpsMGo";

let cachedEmployees: IBreatheHREmployee[] | null = null;
let fetchEmployeesPromise: Promise<IBreatheHREmployee[] | null> | null = null;

export const fetchAllBreatheHREmployees = async (forceRefresh = false): Promise<IBreatheHREmployee[] | null> => {
    if (!forceRefresh && cachedEmployees !== null) {
        return cachedEmployees;
    }

    if (!forceRefresh && fetchEmployeesPromise !== null) {
        return fetchEmployeesPromise;
    }

    fetchEmployeesPromise = (async () => {
        const allEmployees: IBreatheHREmployee[] = [];
    const MAX_PAGES = 100;
    const PER_PAGE = 100;

    try {
        for (let page = 1; page <= MAX_PAGES; page++) {
            try {
                const response = await fetch(
                    `${BREATHEHR_BASE_URL}/employees?page=${page}&per_page=${PER_PAGE}`,
                    {
                        method: 'GET',
                        headers: {
                            Accept: 'application/json',
                            'X-API-KEY': BREATHEHR_API_KEY,
                        },
                    }
                );

                if (!response.ok) {
                    console.warn(`[BreatheHRService] Non-OK response on page ${page}: ${response.status}`);
                    break;
                }

                const data = await response.json();
                const employees: IBreatheHREmployee[] = data?.employees || [];

                if (employees.length === 0) break;

                allEmployees.push(...employees);

                const meta = data?.meta;
                if (meta?.total_pages && page >= meta.total_pages) break;

            } catch (pageErr) {
                console.warn(`[BreatheHRService] Could not fetch employees on page ${page}:`, pageErr);
                break;
            }
        }

        const filtered = allEmployees.filter((e) => e.status === "Current employee");
        console.log(`[BreatheHRService] Total fetched: ${allEmployees.length}, Current employees: ${filtered.length}`);
        cachedEmployees = filtered;
        fetchEmployeesPromise = null;
        return filtered;

    } catch (err) {
        console.error(`[BreatheHRService] Could not fetch BreatheHR employees:`, err);
        fetchEmployeesPromise = null;
        return null;
    }
    })();

    return fetchEmployeesPromise;
};
