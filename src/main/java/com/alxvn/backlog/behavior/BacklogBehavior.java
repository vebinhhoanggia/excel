package com.alxvn.backlog.behavior;

import com.alxvn.backlog.dto.BacklogDetail;
import com.alxvn.backlog.dto.CustomerTarget;

public interface BacklogBehavior {

	public CustomerTarget resolveTarget(BacklogDetail bd);
}
