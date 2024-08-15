revoke usage on schema pgxls from public;
revoke execute on all functions in schema pgxls from public;
revoke execute on all procedures in schema pgxls from public;

grant usage on schema pgxls to :roles;
grant execute on all functions in schema pgxls to :roles;
grant execute on all procedures in schema pgxls to :roles;
