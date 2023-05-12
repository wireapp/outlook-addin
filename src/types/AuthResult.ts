export interface AuthResult {
  success: boolean;
  access_token?: string | null;
  refresh_token?: string | null;
  error?: string | null;
}
