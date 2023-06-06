/* global localStorage */
import { SelfUser } from "../types/SelfUser";

export const setUserDetails = (user: SelfUser) => {
  localStorage.setItem("user", JSON.stringify(user));
};

export const removeUserDetails = () => {
  localStorage.removeItem("user");
};

export const getUserDetails = (): SelfUser | null => {
  const user = localStorage.getItem("user");
  return user ? JSON.parse(user) : null;
};
