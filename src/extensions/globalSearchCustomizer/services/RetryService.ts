export class RetryService {
  public static async withRetry<T>(
    fn: () => Promise<T>,
    maxRetries: number = 3
  ): Promise<T> {
    let attempt = 0;
    while (attempt < maxRetries) {
      try {
        return await fn();
      } catch (error: any) {
        const is429 = 
          error?.statusCode === 429 || 
          error?.status === 429 ||
          error?.message?.includes('429') ||
          error?.message?.toLowerCase().includes('too many requests');
        
        if (is429 && attempt < maxRetries - 1) {
          const delay = Math.pow(2, attempt) * 1000; // 1s, 2s, 4s
          await new Promise(resolve => setTimeout(resolve, delay));
          attempt++;
        } else {
          throw error;
        }
      }
    }
    throw new Error('Max retries reached');
  }
}
