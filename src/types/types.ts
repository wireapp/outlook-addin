export interface AuthResult {
  success: boolean;
  access_token: string;
  refresh_token: string;
}

export interface EventResult {
  id: string;
  link: string;
}
