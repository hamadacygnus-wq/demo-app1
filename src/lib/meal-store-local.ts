import type { MealStore } from "./meal-types";

export const MEAL_STORAGE_KEY = "joy-syokuji-meal-records";

export function loadMealStoreLocal(): MealStore {
  if (typeof window === "undefined") return {};
  try {
    const raw = localStorage.getItem(MEAL_STORAGE_KEY);
    if (!raw) return {};
    const parsed = JSON.parse(raw) as MealStore;
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

export function saveMealStoreLocal(store: MealStore) {
  if (typeof window === "undefined") return;
  localStorage.setItem(MEAL_STORAGE_KEY, JSON.stringify(store));
}
