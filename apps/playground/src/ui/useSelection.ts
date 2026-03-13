import { useCallback, useState } from "react";

export function useSelection(initialSheet = "Sheet1", initialAddr = "A1") {
  const [sheetName, setSheetName] = useState(initialSheet);
  const [address, setAddress] = useState(initialAddr);
  const select = useCallback((nextSheetName: string, nextAddress: string) => {
    setSheetName(nextSheetName);
    setAddress(nextAddress);
  }, []);

  return {
    sheetName,
    address,
    select
  };
}
