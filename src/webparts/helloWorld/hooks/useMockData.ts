import { useEffect, useState } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getDisasters } from '../../../../services';

export function useMockData(context: WebPartContext) {
  const [data, setData] = useState<any>(null);
  const [error, setError] = useState<Error | null>(null);
  const [loading, setLoading] = useState<boolean>(true);

  useEffect(() => {
    let ignore = false;

    setLoading(true);
    setError(null);

    const timeoutId = window.setTimeout(() => {
      const loadData = async (): Promise<void> => {
        try {
          const result = await getDisasters(context);

          if (!ignore) {
            setData(result);
          }
        } catch (err) {
          if (!ignore) {
            setError(
              err instanceof Error
                ? err
                : new Error('Unknown error')
            );
          }
        } finally {
          if (!ignore) {
            setLoading(false);
          }
        }
      };

      void loadData();
    }, 1000);

    return () => {
      ignore = true;
      window.clearTimeout(timeoutId);
    };
  }, [context]);

  return {
    data,
    error,
    loading
  };
}