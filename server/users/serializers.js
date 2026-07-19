export function serializeUser(user) {
  return {
    id: user.id,
    first_name: user.first_name,
    last_name: user.last_name,
    email: user.email,
    name: [user.first_name, user.last_name].filter(Boolean).join(" "),
    timezone: user.timezone
  };
}

export function serializeProfile(profile) {
  return {
    id: profile.id,
    user_id: profile.user_id,
    name: profile.name,
    static_fields: profile.static_fields || {},
    resume_text: profile.resume_text || "",
    preferences: profile.preferences || {},
    profile_version: Number(profile.profile_version || 1),
    created_at: profile.created_at,
    updated_at: profile.updated_at
  };
}
