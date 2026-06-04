/// <reference types="vite/client" />
import type { ExpenseApi } from "./expense-api";

declare global {
  interface Window {
    expenseApi: ExpenseApi;
  }
}
export {};
