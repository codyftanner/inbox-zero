import { redirect } from "next/navigation";
import { NO_LOGIN_REDIRECT_PATH } from "@/utils/redirect";

export default function LoginDisabledLayout() {
  redirect(NO_LOGIN_REDIRECT_PATH);
}
