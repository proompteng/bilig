import { useState } from "react";

export function useSelection(initialSheet = "Sheet1", initialAddr = "A1") {
  const [sheetName, setSheetName] = useState(initialSheet);
  const [address, setAddress] = useState(initialAddr);

  return {
    sheetName,
    address,
    select(nextSheetName: string, nextAddress: string) {
      setSheetName(nextSheetName);
      setAddress(nextAddress);
    }
  };
}
