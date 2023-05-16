import React from "react";
// import { Welcome } from "./sample/Welcome";
import {Suppliers} from "./Suppliers"
import { app } from "@microsoft/teams-js";
import { SuppliersCard } from "./SuppliersCard";
import { useState } from "react";

export default function Tab() {
  const [context, setContext] = useState();

  app.getContext().then((context) => {
    setContext(context);
  });
  
  if (context && context.page.frameContext === "sidePanel") {
    return (
    <div>
    <SuppliersCard />  
   </div>   
    )
  }
  else {
    return (
    <div>
    <Suppliers />  
   </div>
  )
  }
}