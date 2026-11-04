import { createContext, useContext, useEffect, useId } from "react";

export const OperationLock = createContext<(id: string, busy: boolean) => void>(() => {});
export function useOperationLock(busy: boolean) {
  const id = useId();
  const update = useContext(OperationLock);
  useEffect(() => {
    update(id, busy);
    return () => update(id, false);
  }, [id, busy, update]);
}
