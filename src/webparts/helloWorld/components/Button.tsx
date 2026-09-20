import * as React from "react";
import { useState } from "react";

import styles from "./Button.module.scss";

export default function NewButton() {
  const [toggle, setToggle] = useState<boolean>(false);

  return (
    <>
      <button
        className={`${styles.button} ${toggle ? styles.gray : styles.colored}`}
        onClick={() => setToggle(!toggle)}
      >
        Click Me
      </button>
    </>
  );
}
