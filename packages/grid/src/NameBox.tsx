import { useEffect, useState } from "react";
import type { WorkbookDefinedNameSnapshot } from "@bilig/protocol";
import { resolveNameBoxDisplayValue } from "./formulaAssist.js";
import { formulaFieldShellClass, formulaInputClass } from "./formula-bar-theme.js";

interface NameBoxProps {
  readonly address: string;
  readonly definedNames?: readonly WorkbookDefinedNameSnapshot[];
  readonly sheetName: string;
  readonly onCommit: (next: string) => void;
}

export function NameBox({ address, definedNames, sheetName, onCommit }: NameBoxProps) {
  const displayValue = resolveNameBoxDisplayValue({
    sheetName,
    address,
    ...(definedNames ? { definedNames } : {}),
  });
  const [inputValue, setInputValue] = useState(displayValue);

  useEffect(() => {
    setInputValue(displayValue);
  }, [displayValue, sheetName]);

  return (
    <div className="w-[108px] shrink-0">
      <label className="sr-only" htmlFor="name-box-input">
        Name
      </label>
      <div className={formulaFieldShellClass()}>
        <input
          aria-label="Name box"
          className={`${formulaInputClass()} w-full px-2.5 font-medium`}
          data-testid="name-box"
          id="name-box-input"
          value={inputValue}
          onBlur={() => setInputValue(displayValue)}
          onChange={(event) => setInputValue(event.target.value)}
          onKeyDown={(event) => {
            event.stopPropagation();
            if (event.key === "Enter") {
              event.preventDefault();
              onCommit(event.currentTarget.value);
            }
            if (event.key === "Escape") {
              event.preventDefault();
              setInputValue(displayValue);
            }
          }}
        />
      </div>
    </div>
  );
}
