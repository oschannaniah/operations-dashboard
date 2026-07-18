-- Registration now passes first/last name and phone in the auth signUp() call's metadata
-- (LoginScreen used to collect these but the old backend's handleRegister_ wrote them directly
-- to the Users sheet; here they only exist in auth.users.raw_user_meta_data until this trigger
-- copies them onto the new profiles row).
create or replace function handle_new_auth_user() returns trigger as $$
begin
  insert into public.profiles (id, organization_id, email, first_name, last_name, phone, tier)
  values (
    new.id,
    (new.raw_user_meta_data->>'organization_id')::uuid,
    new.email,
    coalesce(new.raw_user_meta_data->>'first_name', ''),
    coalesce(new.raw_user_meta_data->>'last_name', ''),
    coalesce(new.raw_user_meta_data->>'phone', ''),
    'unassigned'
  );
  return new;
end;
$$ language plpgsql security definer;
