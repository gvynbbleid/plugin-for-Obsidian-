<script>
  import { Loader } from '@lucide/svelte';
  import { SquareCheckBig } from '@lucide/svelte';

  let renderStates = $state([]);

  export function initRenderStates(data) {
    data.forEach((param) => {
      renderStates.push({ status: 0, filename: param.file.name });
    });
  }

  export function updateRenderStates(i) {
    renderStates[i].status = 1;
  }
</script>

<div class="progress svelte-export-docx">
  {#each renderStates as item}
    <div>
      {#if item.status}
        <SquareCheckBig size="14" />
      {:else}
        <Loader size="14" />
      {/if}
      {item.filename}
    </div>
  {/each}
</div>

<style>
  :global(.progress.svelte-export-docx) {
    font-size: 14px;
  }
</style>
