import teamsArchiveBackfillStateSchema from '~/schema/teamsArchiveBackfillState';
import type { ITeamsArchiveBackfillState } from '~/schema/teamsArchiveBackfillState';

export function createTeamsArchiveBackfillStateModel(mongoose: typeof import('mongoose')) {
  return (
    mongoose.models.TeamsArchiveBackfillState ||
    mongoose.model<ITeamsArchiveBackfillState>(
      'TeamsArchiveBackfillState',
      teamsArchiveBackfillStateSchema,
    )
  );
}
